# Aba_Dados

Scrapes the **"Aba Dados"** (Data tab) section of **Transferegov** to collect proposal/process data, with an **async** implementation and an optional **e-mail delivery** pipeline.

## 🔁 What it does

- `Aba_Dados_async.py` — async Playwright scraper for the Data tab (fast, concurrent).
- `Aba_Dados_emails.py` — variant that sends collected data via e-mail.
- `Acomp_Fisc_esclarecimentos.py` / `Acomp_Fisc_esclarecimentos_async.py` — acompanhamento & fiscalização ("esclarecimentos") collection (sync + async).

## ▶️ Run

```bash
python Aba_Dados_async.py
```

> Requires a **running Chrome** on debug port `9222` (Playwright CDP), an input `.xlsx` with the target process/proposal numbers (see script header for exact column/sheet names).

## 📦 Dependencies

`playwright` · `pandas` · `aiofiles` · `nest_asyncio` · `tqdm` · `colorama` (and an installed Chrome).