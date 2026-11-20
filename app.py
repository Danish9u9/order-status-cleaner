"""
app.py - the web page. Run locally with:  streamlit run app.py
"""
import hmac
import json
import os

import streamlit as st

from cleaner import STATUS_UNKNOWN, run_report

st.set_page_config(page_title="Order Status Cleaner", page_icon="📦", layout="wide")


def secret(name, default=""):
    """Read a setting from an environment variable first, then from Streamlit secrets."""
    val = os.getenv(name)
    if val:
        return val
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


# ---------------------------------------------------------------
# Optional password gate (set APP_PASSWORD in secrets)
# ---------------------------------------------------------------
APP_PASSWORD = str(secret("APP_PASSWORD", ""))

if APP_PASSWORD and not st.session_state.get("auth"):
    st.title("📦 Order Status Cleaner")
    entered = st.text_input("Password", type="password")
    if entered:
        if hmac.compare_digest(entered.encode(), APP_PASSWORD.encode()):
            st.session_state["auth"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    st.stop()

# ---------------------------------------------------------------
# Main page
# ---------------------------------------------------------------
st.title("📦 Order Status Cleaner")
st.caption("Upload your Shopify orders CSV and any courier shipment CSVs, then click **Generate report**.")

if not APP_PASSWORD:
    st.sidebar.warning("No APP_PASSWORD is set – anyone who has this link can use the app.")

token = str(secret("POSTEX_API_TOKEN", ""))
bulk_url = str(secret("POSTEX_BULK_URL", ""))
single_url = str(secret("POSTEX_SINGLE_URL", ""))

with st.sidebar:
    st.header("Options")
    use_selenium = st.checkbox(
        "Also check courier websites (slow)",
        value=False,
        help="Opens TCS / Leopards / PostEx tracking pages for parcels still 'In Transit' or unknown. "
             "Slow, and may not work on a hosted server.",
    )
    st.markdown("**PostEx API:** " + ("✅ token set" if token else "❌ token missing"))

uploaded = st.file_uploader(
    "CSV files (Shopify orders + shipment files)",
    type=["csv"],
    accept_multiple_files=True,
)

if st.button("Generate report", type="primary", disabled=not uploaded):
    files = [(f.name, f.getvalue()) for f in uploaded]
    try:
        with st.status("Working...", expanded=True) as status:
            result = run_report(
                files,
                token=token,
                bulk_url=bulk_url,
                single_url=single_url,
                use_selenium=use_selenium,
                log=lambda m: st.write(m),
            )
            status.update(label="Done ✅", state="complete", expanded=False)
        st.session_state["result"] = result
    except ValueError as e:
        st.error(str(e))
        st.session_state.pop("result", None)
    except Exception as e:
        st.error(f"Something went wrong: {type(e).__name__}: {e}")
        st.session_state.pop("result", None)

result = st.session_state.get("result")
if result:
    flat = result["flat"]

    if result["unknown"]:
        st.warning(f"{result['unknown']} orders show '{STATUS_UNKNOWN}'. "
                   "Check the PostEx settings, or tick the website-check option.")

    st.download_button(
        "⬇️ Download Excel report",
        data=result["excel"],
        file_name=result["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    counts = flat["Status"].value_counts()
    cols = st.columns(max(len(counts), 1))
    for col, (name, n) in zip(cols, counts.items()):
        col.metric(name, int(n))

    tab1, tab2 = st.tabs(["All orders", "Summary"])
    with tab1:
        wanted = st.multiselect("Filter by status", sorted(flat["Status"].unique()))
        view = flat[flat["Status"].isin(wanted)] if wanted else flat
        st.dataframe(view, hide_index=True)
    with tab2:
        st.dataframe(result["summary"], hide_index=True)

    if result["debug"]:
        with st.expander("Diagnostics (raw PostEx replies – may contain customer data)"):
            st.download_button(
                "Download postex_debug.json",
                data=json.dumps(result["debug"], indent=2, ensure_ascii=False),
                file_name="postex_debug.json",
                mime="application/json",
            )
