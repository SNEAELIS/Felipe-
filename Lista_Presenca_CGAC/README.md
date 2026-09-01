# Lista_Presenca_CGAC

Computes **attendance-list statistics** ("lista de presença") for CGAC meetings from Excel sheets containing presence/fault marks.

## 🔁 What it does

- `lista_presenca_cgac.py` — reads `.xlsx` attendance files and calculates per-member/per-meeting stats (`P` = presente, `F` = falta, averages, and a formatted output workbook..

## ▶️ Run

```bash
python lista_presenca_cgac.py
```

> Point the script at a folder containing the attendance `.xlsx` files (see script header for layout).

## 📦 Dependencies

`pandas` · `numpy` · `openpyxl`