# Pareceres_Emendas_PIX

Automates the extraction/processing of **"Emendas PIX" legal opinions** ("pareceres") — the workflow that pairs federal budget amendments(emendas) with PIX-based transfers..

## 🔁 What it does

- `Pareceres_Emendas_PIX.py` — Playwright automation that processes emendas/PIX opinion data(searches, extracts, registers) for the involved processes; see script header for input format..

## ▶️ Run

```bash
python Pareceres_Emendas_PIX.py
```

> Requires a **running Chrome** on debug port(Playwright CDP) + portal access limited to authorized users..

## 📦 Dependencies

`playwright` · `pandas` · `openpyxl` · `colorama`