# Acompanhamento

Automates **follow-up / monitoring** ("acompanhamento") workflows on **Transferegov**, scraping progress and fiscalization data for processes/proposals.

## 🔁 What it does

- `Acompanhamento.py` — Selenium-based scraper for acompanhamento data on the portal..

## ▶️ Run

```bash
python Acompanhamento.py
```

> Requires an input `.xlsx` with the target **process numbers** and Chrome available (see script header for the exact sheet/column).

## 📦 Dependencies

`selenium` · `webdriver-manager` · `pandas` · `openpyxl` · `tqdm` · `colorama`