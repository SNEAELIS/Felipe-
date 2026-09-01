# CAUC_Phone_plus_email

Looks up **CAUC clearance data** for instruments/processes and extracts **phone + e-mail** contact information automatically.

## 🔁 What it does

- `CAUC_Phone_plus_email.py` — Selenium automation that navigates the portal, fetches CAUC status, and collects contact data for each item in the input spreadsheet..

## ▶️ Run

```bash
python CAUC_Phone_plus_email.py
```

> The script consumes an `.xlsx` file listing the items to consult and writes results back to an output spreadsheet (names in the script header). Chrome required.



## 📦 Dependencies

`selenium` · `webdriver-manager` · `pandas` · `openpyxl` · `colorama`