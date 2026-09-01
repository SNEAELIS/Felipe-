# Consultas_DFP

Automates **DFP consultations** on **Transferegov** — collecting transfer/data information for the DFP (Diretoria de Fomento e Programas) workflows.

## 🔁 What it does

- `Consultas_DFP_transf_gov.py` — Selenium scraper that navigates Transferegov consultations for a list of processes/proposals and gathers the relevant DFP data into an `.xlsx` output.



## ▶️ Run

```bash
python Consultas_DFP_transf_gov.py
```

> Input `.xlsx` format: see script header. Chrome + internet required.





## 📦 Dependencies

`selenium` · `webdriver-manager` · `pandas` · `openpyxl` · `numpy` · `pytesseract` · `pyautogui` · `autoit`