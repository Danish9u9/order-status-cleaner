"""
cleaner.py - the core logic of the Order Status Cleaner (no user interface here).
app.py calls run_report(); you normally never need to edit this file.
"""
import io
import re
import shutil
import time
import warnings

import openpyxl
import pandas as pd
import requests
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BIG_ORDER_LIMIT = 9000
POSTEX_BATCH_SIZE = 20
STATUS_UNKNOWN = "Check Manually"


# =============================================================
# STATUS NORMALISER (one place, used for every source)
# =============================================================
def normalize_status(raw):
    """Map any courier wording to: Delivered / Returned / Cancelled / Rebooked / In Transit."""
    s = str(raw or "").strip().lower()
    if not s or s == "nan":
        return "In Transit"
    if "rebook" in s:
        return "Rebooked"                       # old booking that was replaced -> ignore
    # failures/returns FIRST, otherwise "undelivered" would match "delivered"
    if re.search(r"\b(un-?delivered|not delivered|delivery failed)\b", s) or \
       re.search(r"return|\brto\b|refused|back to shipper", s):
        return "Returned"
    if "cancel" in s:
        return "Cancelled"
    if re.search(r"\bdelivered\b", s):
        return "Delivered"
    return "In Transit"


# =============================================================
# TRACKING NUMBERS
# =============================================================
TRACK_PATTERNS = [
    (re.compile(r"\bKI\d{8,12}\b", re.I), "Leopards"),
    (re.compile(r"\b173\d{9}\b"), "TCS"),
    (re.compile(r"\b\d{14}\b"), "PostEx"),          # PostEx CNs are 14 digits
]


def courier_of(tracking):
    for pat, name in TRACK_PATTERNS:
        if pat.fullmatch(tracking):
            return name
    return "Unknown"


def _text(v):
    return "" if pd.isna(v) else str(v)


def extract_trackings(row):
    text = f"{_text(row.get('Note Attributes'))} {_text(row.get('Notes'))}"
    found = []
    for pat, _ in TRACK_PATTERNS:
        for m in pat.findall(text):
            m = m.upper()
            if m not in found:
                found.append(m)
    return found


def parse_price(val):
    if pd.isna(val):
        return 0.0
    s = str(val).replace(",", "").replace("Rs.", "").replace("PKR", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


# =============================================================
# LOADING FILES
# =============================================================
def classify_files(files, log=print):
    """files = [(filename, bytes)] -> (orders_file or None, [shipment_files])"""
    orders, shipments = None, []
    for name, data in files:
        try:
            cols = [str(c).lower() for c in pd.read_csv(io.BytesIO(data), nrows=2).columns]
        except Exception as e:
            log(f"⚠️ Could not read {name}: {e}")
            continue
        if "tracking number" in cols or "reference no" in cols:
            log(f"   -> Shipment/Tracking file: {name}")
            shipments.append((name, data))
        elif "name" in cols and ("tags" in cols or "total" in cols):
            log(f"   -> Shopify Orders file:    {name}")
            orders = (name, data)
        else:
            log(f"   -> Ignored (not orders/shipments): {name}")
    return orders, shipments


def load_orders(data):
    df = pd.read_csv(io.BytesIO(data), low_memory=False)
    order_level_cols = ["Financial Status", "Total", "Shipping", "Notes", "Note Attributes",
                        "Tags", "Shipping Name", "Billing Name", "Created at", "Cancelled at"]
    existing = [c for c in order_level_cols if c in df.columns]
    if existing:
        df[existing] = df.groupby("Name")[existing].ffill()
    df = df.drop_duplicates(subset=["Name"]).reset_index(drop=True)
    df["Price"] = df["Total"].apply(parse_price) if "Total" in df.columns else 0.0
    return df


def _booked(row, sort_cols):
    """Booking date/time as a real timestamp so 'newest booking' is truly the newest."""
    parts = [str(row[c]).strip() for c in sort_cols if pd.notna(row[c])]
    if not parts:
        return pd.Timestamp.min
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        ts = pd.to_datetime(" ".join(parts), errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return pd.Timestamp.min
    if ts.tzinfo is not None:
        ts = ts.tz_convert(None)
    return ts


def load_shipments(shipment_files):
    out = {}                # order_id -> list of parcel dicts
    for _, data in shipment_files:
        df = pd.read_csv(io.BytesIO(data), dtype=str)
        cols = {str(c).lower(): c for c in df.columns}
        ref_col, track_col, status_col = cols.get("reference no"), cols.get("tracking number"), cols.get("status")
        courier_col = cols.get("courier")
        if not (ref_col and track_col):
            continue
        sort_cols = [c for c in (cols.get("booking date"), cols.get("booking time")) if c]
        for _, r in df.iterrows():
            if pd.isna(r[ref_col]) or pd.isna(r[track_col]):
                continue
            oid = str(r[ref_col]).strip().replace("#", "")
            raw = r[status_col] if status_col and pd.notna(r[status_col]) else ""
            out.setdefault(oid, []).append({
                "tracking": str(r[track_col]).strip().upper(),
                "raw": raw,
                "status": normalize_status(raw),
                "courier": (r[courier_col] if courier_col and pd.notna(r[courier_col]) else "Insta"),
                "booked": _booked(r, sort_cols),
            })
    for lst in out.values():
        lst.sort(key=lambda x: x["booked"])             # oldest -> newest booking
    return out


def build_order_parcels(df_main, csv_shipments):
    """Every parcel per order: CSV parcels + parcels found in notes (merged, de-duplicated)."""
    order_parcels = {}
    for _, row in df_main.iterrows():
        oid = str(row["Name"]).strip().replace("#", "")
        parcels = [dict(p) for p in csv_shipments.get(oid, [])]
        known = {p["tracking"] for p in parcels}
        for t in extract_trackings(row):
            if t not in known:
                parcels.append({"tracking": t, "raw": "", "status": None,
                                "courier": courier_of(t), "booked": pd.Timestamp.min})
        for p in parcels:
            c = courier_of(p["tracking"])
            p["courier"] = c if c != "Unknown" else p["courier"]
        order_parcels[oid] = parcels
    return order_parcels


# =============================================================
# POSTEX LIVE STATUS (bulk first, single-track fallback)
# =============================================================
def _postex_status_from_item(item):
    resp = item.get("trackingResponse", item)
    if isinstance(resp, list):
        resp = resp[0] if resp else {}
    if not isinstance(resp, dict):
        return ""
    raw = resp.get("transactionStatus") or resp.get("orderStatus") or resp.get("status") or ""
    if raw:
        return str(raw)
    hist = resp.get("transactionStatusHistory") or []
    msgs = [str(h.get("transactionStatusMessage", "")) for h in hist if isinstance(h, dict)]
    for wanted in ("Delivered", "Returned", "Cancelled"):
        for m in msgs:
            if normalize_status(m) == wanted:
                return m
    return msgs[-1] if msgs else ""


def _get(session, url, log, **kw):
    for attempt in range(3):
        try:
            r = session.get(url, timeout=20, **kw)
            if r.status_code == 200:
                return r
            log(f"   ⚠️ PostEx HTTP {r.status_code}")
            if r.status_code in (401, 403):      # bad token - retrying won't help
                return None
        except requests.RequestException as e:
            log(f"   ⚠️ PostEx request error (try {attempt + 1}/3): {type(e).__name__}")
        time.sleep(1.5 * (attempt + 1))
    return None


def fetch_postex_statuses(tracking_numbers, token, bulk_url, single_url, log, debug):
    result = {}
    trackings = sorted({t for t in tracking_numbers if t})
    if not trackings:
        return result
    token = (token or "").strip()
    if not token:
        log("❌ No PostEx API token set – PostEx orders will show 'Check Manually'.")
        return result
    bulk_ok = str(bulk_url or "").startswith("http")
    single_ok = str(single_url or "").startswith("http") and "{}" in single_url
    if not (bulk_ok or single_ok):
        log("❌ PostEx API URLs are not set (POSTEX_BULK_URL / POSTEX_SINGLE_URL) – skipping PostEx.")
        return result

    session = requests.Session()
    session.headers.update({"token": token, "Content-Type": "application/json"})

    for i in range(0, len(trackings), POSTEX_BATCH_SIZE):
        batch = trackings[i:i + POSTEX_BATCH_SIZE]
        log(f"Fetching {len(batch)} PostEx trackings (batch {i // POSTEX_BATCH_SIZE + 1})...")
        if bulk_ok:
            r = _get(session, bulk_url, log, params={"TrackingNumbers": ",".join(batch)})
            if r is not None:
                try:
                    data = r.json()
                    debug[f"bulk_{i}"] = data
                    dist = data.get("dist") or []
                    if isinstance(dist, dict):
                        dist = [dist]
                    for item in dist:
                        tn = str(item.get("trackingNumber", "")).strip()
                        raw = _postex_status_from_item(item)
                        if tn and raw:
                            result[tn] = raw
                except ValueError:
                    log("   ⚠️ PostEx returned non-JSON for this batch.")
        if single_ok:
            for tn in [t for t in batch if t not in result]:
                r1 = _get(session, single_url.format(tn), log)
                if r1 is None:
                    continue
                try:
                    data = r1.json()
                    debug[f"single_{tn}"] = data
                    item = data.get("dist") or {}
                    if isinstance(item, list):
                        item = item[0] if item else {}
                    raw = _postex_status_from_item(item) if isinstance(item, dict) else ""
                    if raw:
                        result[tn] = raw
                except ValueError:
                    pass

    missing = [t for t in trackings if t not in result]
    log(f"   PostEx answered {len(result)}/{len(trackings)} trackings; {len(missing)} unresolved.")
    return result


# =============================================================
# DECIDE ONE STATUS PER ORDER
# =============================================================
def resolve_order(parcels):
    """-> (status, parcels_to_display, raw_status_text)"""
    if not parcels:
        return None, [], ""
    known = [p for p in parcels if p["status"]]
    unknown = [p for p in parcels if not p["status"]]

    delivered = [p for p in known if p["status"] == "Delivered"]
    if delivered:                                               # delivered always wins
        return "Delivered", [delivered[0]], delivered[0]["raw"]

    live = [p for p in known if p["status"] not in ("Rebooked", "Cancelled")]
    if live:
        last = live[-1]                                         # newest active booking
        return last["status"], parcels, last["raw"]
    if unknown:                                                 # tracking exists but nobody could tell us its status
        return STATUS_UNKNOWN, parcels, ""
    if any(p["status"] == "Cancelled" for p in known):
        return "Cancelled", parcels, known[-1]["raw"]
    return "In Transit", parcels, known[-1]["raw"]              # only 'Rebooked' left


# =============================================================
# OPTIONAL: SELENIUM FALLBACK (courier websites)
# =============================================================
WEB_URLS = {
    "PostEx":   "https://postex.pk/tracking?cn={}",
    "TCS":      "https://www.tcsexpress.com/track/{}",
    "Leopards": "https://pk.leopardscourier.com/tracking",
}
WEB_IGNORE = re.compile(r"policy|estimated|delivery time|how to|track your|enter your|terms|contact us|faq|charges", re.I)


def web_status_from_text(body_text):
    """Read a tracking page's visible text line by line. Returns (status, evidence_line) or (None, "")."""
    hits = {"Delivered": None, "Returned": None, "Cancelled": None, "In Transit": None}
    for line in (l.strip() for l in body_text.splitlines()):
        if not line or len(line) > 120 or WEB_IGNORE.search(line):
            continue
        st = normalize_status(line)
        if st == "In Transit":
            if re.search(r"in[- ]?transit|out for delivery|dispatched|on the way|at .*(hub|warehouse|facility)|booked|picked", line, re.I):
                hits["In Transit"] = hits["In Transit"] or line
            continue
        if st in hits:
            hits[st] = hits[st] or line
    for st in ("Delivered", "Returned", "Cancelled", "In Transit"):
        if hits[st]:
            return st, hits[st]
    return None, ""


def _make_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    opts = Options()
    for a in ("--headless=new", "--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage", "--window-size=1920,1080"):
        opts.add_argument(a)
    chromium = shutil.which("chromium") or shutil.which("chromium-browser")
    chromedriver = shutil.which("chromedriver")
    if chromium and chromedriver:                    # Linux server / Streamlit Cloud
        opts.binary_location = chromium
        return webdriver.Chrome(service=Service(chromedriver), options=opts)
    from webdriver_manager.chrome import ChromeDriverManager   # Windows / Mac with Chrome installed
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)


def selenium_recheck(order_parcels, skip_orders, log):
    targets = []
    for oid, ps in order_parcels.items():
        if oid in skip_orders:
            continue
        st = resolve_order(ps)[0]
        if st == STATUS_UNKNOWN:
            parcels = [p for p in ps if not p["status"]]
        elif st == "In Transit":
            parcels = [p for p in ps if p["status"] == "In Transit"]
        else:
            continue
        targets += [(oid, p) for p in parcels if p["courier"] in WEB_URLS]
    if not targets:
        return
    log(f"🌐 {len(targets)} parcels are 'In Transit' or unknown – checking courier websites...")
    try:
        from selenium.webdriver.common.by import By
        driver = _make_driver()
    except Exception as e:
        log(f"⚠️ Could not start Chrome/Selenium ({type(e).__name__}). Those orders stay '{STATUS_UNKNOWN}'.")
        return

    def page_text():
        try:
            return driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            return ""

    try:
        for n, (oid, p) in enumerate(targets, 1):
            t, courier = p["tracking"], p["courier"]
            try:
                if courier == "Leopards":
                    driver.get(WEB_URLS[courier])
                    time.sleep(2)
                    box = driver.find_element(By.XPATH, "//input[contains(@name,'track') or contains(@placeholder,'rack') or contains(@id,'track')]")
                    box.clear()
                    box.send_keys(t)
                    driver.find_element(By.XPATH, "//button[@type='submit' or contains(.,'Track')]").click()
                else:
                    driver.get(WEB_URLS[courier].format(t))
                status, evidence = None, ""
                for _ in range(8):
                    time.sleep(2)
                    text = page_text()
                    if t.lower() in text.lower() or courier == "PostEx":
                        status, evidence = web_status_from_text(text)
                        if status in ("Delivered", "Returned", "Cancelled"):
                            break
                if status and (not p["status"] or status in ("Delivered", "Returned", "Cancelled")):
                    p["status"], p["raw"] = status, f"Website ({courier}): {evidence}"
                log(f"   [{n}/{len(targets)}] #{oid} {courier} {t} -> {status or 'no status found'}")
            except Exception as e:
                log(f"   [{n}/{len(targets)}] #{oid} {courier} {t} -> could not check ({type(e).__name__})")
    finally:
        driver.quit()


# =============================================================
# BUILD REPORT TABLES
# =============================================================
def build_frames(df_main, order_parcels):
    keys = ["simple", "big", "self", "reattempt", "bycall"]
    buckets = {k: {"orders": [], "statuses": []} for k in keys}
    flat_rows = []

    for _, row in df_main.iterrows():
        raw_id = str(row["Name"]).strip()
        oid = raw_id.replace("#", "")
        disp_id = raw_id if raw_id.startswith("#") else f"#{raw_id}"
        tags = str(row.get("Tags", "")).lower() if pd.notna(row.get("Tags", "")) else ""
        price = row["Price"]

        is_reattempt = any(k in tags for k in ["re-attempt by danish", "reattempt by danish", "re attempt by danish"])
        is_self = "self delivery" in tags
        is_by_call = "by call" in tags
        is_big = price > BIG_ORDER_LIMIT or "big order" in tags
        shopify_cancelled = pd.notna(row.get("Cancelled at", None))

        parcels = order_parcels.get(oid, [])
        status, shown, raw_status = resolve_order(parcels)

        if is_self:
            status, shown, raw_status = "Delivered", parcels, "Self Delivery"
        elif status is None:
            status = "Cancelled" if shopify_cancelled else "No Tracking ID"
            shown = []

        tracking_part = f" [Trackings: {', '.join(p['tracking'] for p in shown)}]" if shown else ""
        if is_big:
            display = f"{disp_id} (Big Order){tracking_part}"
        elif not shown:
            display = f"{disp_id} (No Tracking ID)"
        else:
            display = f"{disp_id}{tracking_part}"

        if is_by_call:
            key, display = "bycall", disp_id
        elif is_reattempt:
            key = "reattempt"
        elif is_big:
            key = "big"
        elif is_self:
            key = "self"
        else:
            key = "simple"
        buckets[key]["orders"].append(display)
        buckets[key]["statuses"].append(status)

        flat_rows.append({
            "Order": disp_id,
            "Category": {"simple": "Simple", "big": "Big Order", "self": "Self Delivery",
                         "reattempt": "Re-Attempt by Danish", "bycall": "By Call"}[key],
            "Courier": ", ".join(sorted({p["courier"] for p in shown})) if shown else "",
            "Tracking": ", ".join(p["tracking"] for p in shown),
            "Status": status,
            "Courier's Own Status": raw_status,
            "Total": price,
        })

    def pad(lst, n):
        return lst + [""] * (n - len(lst))

    max_len = max([len(b["orders"]) for b in buckets.values()] + [1])
    final_df = pd.DataFrame({
        "Simple Orders": pad(buckets["simple"]["orders"], max_len),
        "Simple Order Status": pad(buckets["simple"]["statuses"], max_len),
        "Big Orders (>9k)": pad(buckets["big"]["orders"], max_len),
        "Big Order Status": pad(buckets["big"]["statuses"], max_len),
        "Self Delivery": pad(buckets["self"]["orders"], max_len),
        "Self Delivery Status": pad(buckets["self"]["statuses"], max_len),
        "Re-Attempt by Danish": pad(buckets["reattempt"]["orders"], max_len),
        "Re-Attempt Status": pad(buckets["reattempt"]["statuses"], max_len),
        "EmptyCol": [""] * max_len,
        "By Call": pad(buckets["bycall"]["orders"], max_len),
        "By Call Status": pad(buckets["bycall"]["statuses"], max_len),
    })
    flat_df = pd.DataFrame(flat_rows)
    if flat_df.empty:
        summary_df = pd.DataFrame(columns=["Category", "Total"])
    else:
        summary_df = (flat_df.pivot_table(index="Category", columns="Status", values="Order",
                                          aggfunc="count", fill_value=0)
                      .assign(Total=lambda d: d.sum(axis=1)).reset_index())
        summary_df.columns.name = None
    return final_df, flat_df, summary_df


# =============================================================
# EXCEL OUTPUT + FORMATTING
# =============================================================
def build_excel(final_df, flat_df, summary_df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Order Status Report", index=False)
        flat_df.to_excel(writer, sheet_name="All Orders (flat)", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    big_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    reattempt_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    status_style = {
        "Delivered":      ("E2EFDA", "375623"),
        "Returned":       ("FCE4D6", "C65911"),
        "Cancelled":      ("FFF2CC", "7F6000"),
        "In Transit":     ("DEEAF6", "1F4E78"),
        STATUS_UNKNOWN:   ("D9D9D9", "404040"),
        "No Tracking ID": ("EDEDED", "7F7F7F"),
    }
    border_thin = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

    def style_status(cell):
        v = str(cell.value or "")
        if v in status_style:
            bg, fg = status_style[v]
            cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            cell.font = Font(color=fg, bold=True)

    def autosize(ws, skip=()):
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            if letter in skip:
                ws.column_dimensions[letter].width = 4
                continue
            longest = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[letter].width = min(max(longest + 4, 14), 70)

    ws = wb["Order Status Report"]
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        if cell.value == "EmptyCol":
            cell.value = ""
            continue
        cell.fill, cell.font = header_fill, header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            header = str(ws.cell(row=1, column=c).value or "")
            if header == "" and c == 9:
                continue
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if cell.value:
                if "Big Orders" in header:
                    cell.fill = big_fill
                elif "Re-Attempt by Danish" in header:
                    cell.fill = reattempt_fill
            if "Status" in header:
                style_status(cell)
    autosize(ws, skip={"I"})
    ws.freeze_panes = "A2"

    for name in ("All Orders (flat)", "Summary"):
        w = wb[name]
        for c in range(1, w.max_column + 1):
            h = w.cell(row=1, column=c)
            h.fill, h.font = header_fill, header_font
            h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in range(2, w.max_row + 1):
            for c in range(1, w.max_column + 1):
                cell = w.cell(row=r, column=c)
                cell.border = border_thin
                if str(w.cell(row=1, column=c).value) == "Status":
                    style_status(cell)
        autosize(w)
        w.freeze_panes = "A2"
    wb["All Orders (flat)"].auto_filter.ref = wb["All Orders (flat)"].dimensions

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =============================================================
# MAIN ENTRY POINT
# =============================================================
def run_report(files, token="", bulk_url="", single_url="", use_selenium=False, log=print):
    """
    files: list of (filename, bytes) - the uploaded CSVs (Shopify orders + any shipment files)
    Returns dict: filename, excel (bytes), flat (DataFrame), summary (DataFrame),
                  unknown (int), debug (dict of raw PostEx replies)
    """
    log("Reading uploaded files...")
    orders_file, shipment_files = classify_files(files, log)
    if not orders_file:
        raise ValueError("Could not find a Shopify orders CSV among the uploaded files "
                         "(it needs 'Name' and 'Tags' or 'Total' columns).")

    log("Loading Shopify orders...")
    df_main = load_orders(orders_file[1])
    csv_shipments = load_shipments(shipment_files)
    order_parcels = build_order_parcels(df_main, csv_shipments)

    if "Created at" in df_main.columns:
        months = pd.to_datetime(df_main["Created at"], errors="coerce", utc=True).dt.strftime("%B").dropna()
        month_name = months.mode()[0] if not months.empty else "Report"
    else:
        month_name = "Report"

    debug = {}
    postex_nums = [p["tracking"] for ps in order_parcels.values() for p in ps if p["courier"] == "PostEx"]
    postex_raw = fetch_postex_statuses(postex_nums, token, bulk_url, single_url, log, debug)
    for ps in order_parcels.values():
        for p in ps:
            if p["courier"] == "PostEx" and p["tracking"] in postex_raw:
                p["raw"] = postex_raw[p["tracking"]]
                p["status"] = normalize_status(p["raw"])

    if use_selenium:
        self_orders = {str(r["Name"]).strip().replace("#", "") for _, r in df_main.iterrows()
                       if "self delivery" in str(r.get("Tags", "")).lower()}
        selenium_recheck(order_parcels, self_orders, log)

    log("Processing orders and mapping statuses...")
    final_df, flat_df, summary_df = build_frames(df_main, order_parcels)
    excel = build_excel(final_df, flat_df, summary_df)
    unknown = int((flat_df["Status"] == STATUS_UNKNOWN).sum()) if not flat_df.empty else 0

    return {
        "filename": f"Danish, Bonus Status, {month_name}.xlsx",
        "excel": excel,
        "flat": flat_df,
        "summary": summary_df,
        "unknown": unknown,
        "debug": debug,
    }
