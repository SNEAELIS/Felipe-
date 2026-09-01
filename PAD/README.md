# PAD

Automates extraction of **PAD** (Programa de Ação Detalhado / Plan data) information from **Transferegov**.

## 🔁 What it does

- `PAD.py` — Selenium scraper that reads an input `.xlsx` of processes/proposals and collects the PAD-related data/attachments into organized per-process output folders..

## ▶️ Run

```bash
python PAD.py
```

> Input `.xlsx` with the target numbers; Chrome + portal access required (see script header for column/sheet).

## 📦 Dependencies

`selenium` · `webdriver-manager` · `pandas` · `openpyxl` · `colorama`