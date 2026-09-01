# SEi_SAA

**SEI dashboard** suite for the SAA workflow — multiple entry points (web dashboard, headless, and SOAP API-based).

## 🔁 What it does

- `SEi_dashboard.py` — main SEI dashboard (Playwright-driven data collection + reporting).
- `SEi_dashboard_headless.py` — headless variant for scheduler-friendly runs..
- `sei_dashboard_api.py` — **API-based version** that talks to SEI's **SOAP Web Services** (via `zeep`) instead of browser automation — requires `sei_credentials.json` (gitignored) with URL/chave/IDs.
- `rps_saa.py` / `rps_saa_backup.py` — supporting SAA scrapers (with backup variant).

## ▶️ Run

```bash
python SEi_dashboard.py            # web dashboard
python sei_dashboard_api.py        # SOAP API version (needs sei_credentials.json)
```

> Credentials are **never committed**: create local `sei_credentials.json` (gitignored) from the template documented in the script header..

## 📦 Dependencies

`playwright` · `pandas` · `openpyxl` · `pdfplumber` · `zeep` (API version) · `streamlit` (dashboard UI)