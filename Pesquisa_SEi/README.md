# Pesquisa_SEi

Automates **SEI searches** — querying processes on SEI (Sistema Eletrônico de Informações) in bulk.

## 🔁 What it does

- `Pesquisa_SEi_1-2.py` / `Pesquisa_SEi_2-2.py` — Selenium automation that searches process numbers on SEI e extracts results (split into two parts)..
- `run_pesq_sei.bat` — convenience launcher (kept local, not tracked).

## ▶️ Run

```bash
python Pesquisa_SEi_1-2.py
```

> SEI access/credentials: authorized portal user required. Adjust input paths in-the script header.

## 📦 Dependencies

`selenium` · `webdriver-manager` · `pandas` · `openpyxl` · `colorama`