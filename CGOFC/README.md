# CGOFC

Budget/execution **spreadsheet automation** for the CGOFC area — updates `.xlsx` workbooks programmatically (with a `tkinter` desktop UI for the standalone version).

## 🔁 What it does

- `autom_xlsx_update.py` — automatic updates to budget/execution `.xlsx` files..
- `financeiro.py` — financial/execution data processing helper..
- `stand_alone_auto_xlsx_upload.py` — standalone `tkinter` app to pick a workbook and apply the automated updates..

## ▶️ Run (standalone UI)

```bash
python stand_alone_auto_xlsx_upload.py
```

> Headless/batch use: `python financeiro.py` (see script headers for input formats).

## 📦 Dependencies

`pandas` · `openpyxl`· (`tkinter` ships with Python on Windows)