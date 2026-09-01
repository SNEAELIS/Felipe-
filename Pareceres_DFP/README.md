# Pareceres_DFP

Automates the collection/processing of **DFP legal opinions** ("pareceres") on federal workflows.

## 🔁 What it does

- `parecere_dfp.py` — Playwright automation that queries processes and extracts/registers parecer data (see script header for input details).

## ▶️ Run

```bash
python parecere_dfp.py
```

> Requires a **running Chrome** on debug port (Playwright CDP) on the machine.

## 📦 Dependencies

`playwright` · `pandas` · `openpyxl` · `colorama`