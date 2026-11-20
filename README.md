# Order Status Cleaner (web version)

Upload CSVs in the browser, click one button, download the Excel report.

## Files in this folder

| File | What it is | Upload to GitHub? |
|---|---|---|
| `app.py` | The web page | Yes |
| `cleaner.py` | Your logic (same rules as your original script) | Yes |
| `requirements.txt` | Python packages | Yes |
| `packages.txt` | Chrome for the optional website check (delete if you never use it) | Yes |
| `.gitignore` | Stops secrets/data being uploaded | Yes |
| `run.bat` | Double-click launcher for Windows (local use) | Optional |
| `.streamlit/secrets.toml.example` | Template for your token/password | Template only |
| `.streamlit/secrets.toml` | **Your real secrets (you create this)** | **NEVER** |

## Before anything: fill in the PostEx URLs

Open PostEx's API documentation and copy the bulk-tracking URL into `POSTEX_BULK_URL`
and the single-tracking URL into `POSTEX_SINGLE_URL` (write `{}` where the tracking number goes).
Without them, PostEx orders show "Check Manually".

## Option A – Hosted URL (Streamlit Community Cloud, free)

1. Create a free account at github.com -> **New repository** -> name it `order-cleaner`, set it to **Private**.
2. Click **Add file -> Upload files** and upload: `app.py`, `cleaner.py`, `requirements.txt`,
   `packages.txt`, `.gitignore`. (Do NOT upload any secrets or CSV files.)
3. Go to share.streamlit.io, sign in with GitHub, click **Create app**.
   Pick your repository, branch `main`, main file `app.py`.
4. Open **Advanced settings -> Secrets** and paste (with your real values):
   ```
   POSTEX_API_TOKEN = "your-token"
   POSTEX_BULK_URL = "https://..."
   POSTEX_SINGLE_URL = "https://.../{}"
   APP_PASSWORD = "a-strong-password"
   ```
5. Click **Deploy**. After a few minutes you get a link like
   `https://your-name-order-cleaner.streamlit.app`. Bookmark it and share it with your team.

Notes:
- Always set `APP_PASSWORD`. Apps are reachable by anyone with the link, and the reports contain customer data.
- Free apps go to sleep when unused. Opening the link wakes them in under a minute.
- To change a secret later: App settings -> Secrets.
- To update the code: edit the file on GitHub; the app redeploys automatically.

## Option B – On your own PC / office network

1. Install Python 3.10+ (tick "Add Python to PATH").
2. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill it in.
3. Double-click `run.bat`. The page opens at http://localhost:8501.
4. Coworkers on the same Wi-Fi/office network can open the "Network URL" that the black window shows
   (looks like `http://192.168.x.x:8501`) while your PC is on.

Option B keeps customer data on your own machine, and the Chrome website check works best here.

## The "Also check courier websites" option

It works on Option B (needs Google Chrome installed). On the hosted version it may be slow,
may fail, or courier sites may block the server, so leave it off there.
