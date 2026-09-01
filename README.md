# SNEAELIS — Federal Service Automation Suite

<div align="center">

**Python automation suite for Brazilian federal government workflows** — web scraping, data extraction, dashboards, document collection, and reporting around **Transferegov**, **SEI**, and related public systems.

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green)
![Status](https://img.shields.io/badge/status-active-brightgreen)
![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen)

</div>

---

## 📌 Purpose

This repository centralizes the **automation toolkit used internally by the SNEAELIS team** (Ministry of Social Development / federal environment) for:

- Collecting public data from the Brazilian government transparency portal **Transferegov** (proposals, contracts, payment settlements, accountability, special transfers).
- Interacting with **SEI** (Sistema Eletrônico de Informações) — dashboards, document extraction, and SOAP Web Services.
- Office productivity automations: attendance lists, e-mail dispatch, MESP payroll double-checks.
- Producing clean, inspectable reports in **Excel (.xlsx)**.

The suite is deliberately **modular** — each folder is an independent tool with its own README, dependencies, and purpose.

> ⚠️ **Official-use tooling, not a library.** Every script performs live automation (Selenium / Playwright) against government portals and consumes/updates `.xlsx` files. It is intended for **authorized federal users only**, and all operations must comply with applicable data-protection rules (LGPD, institutional norms).

---

## 🗂️ Repository Structure
| Folder | Purpose | Tech |
|---|---|---|
| [`Aba_Dados/`](Aba_Dados/) | "Aba Dados" extraction from Transferegov — async + e-mail delivery | Playwright (async) |
| [`Acompanhamento/`](Acompanhamento/) | Monitoring / follow-up scraping on Transferegov | Selenium |
| [`Analise_custos/`](Analise_custos/) | Cost-analysis document download per proposal | Selenium |
| [`Analise_Custos_Exec_plus_Print/`](Analise_Custos_Exec_plus_Print/) | Execution + fiscalization print (PDF) per proposal | Selenium + PyAutoGUI |
| [`CAUC_Phone_plus_email/`](CAUC_Phone_plus_email/) | CAUC clearance lookups — phone + e-mail extraction | Selenium |
| [`CGAC_Contratos/`](CGAC_Contratos/) | Download contracts / sub-agreements from Transferegov | Selenium |
| [`CGAC_Doc_Liquidacao/`](CGAC_Doc_Liquidacao/) | Settlement ("Liquidação") documents download | Selenium |
| [`CGAC_Presta_Contas/`](CGAC_Presta_Contas/) | Accountability ("Prestação de Contas") documents | Selenium |
| [`CGOFC/`](CGOFC/) | Budget/execution spreadsheet automation (tkinter UI) | openpyxl |
| [`Consultas_DFP/`](Consultas_DFP/) | DFP consultations on Transferegov (transferências) | Selenium |
| [`emails_poli/`](emails_poli/) | Bulk e-mail sender with domain auto-correction | Outlook COM |
| [`Lista_Presenca_CGAC/`](Lista_Presenca_CGAC/) | Attendance-list statistical analysis | pandas/openpyxl |
| [`PAD/`](PAD/) | PAD (Plano de Ação Detalhado) extraction | Selenium |
| [`Pareceres_DFP/`](Pareceres_DFP/) | DFP legal opinions ("pareceres") automation | Playwright |
| [`Pareceres_Emendas_PIX/`](Pareceres_Emendas_PIX/) | "Emendas PIX" opinions automation | Playwright |
| [`Pesquisa_SEi/`](Pesquisa_SEi/) | SEI search automation | Selenium |
| [`SEi_SAA/`](SEi_SAA/) | SEI dashboards (web + headless + SOAP API) | Playwright + Streamlit |
| [`SEi_SNEAELIS/`](SEi_SNEAELIS/) | SEI dashboard (SNEAELIS-specific) | Playwright |
| [`Selecao_PAC_DIE/`](Selecao_PAC_DIE/) | PAC selection module — docs + map screenshots | Selenium + PIL |
| [`Tel+E_mail_DPF/`](Tel+E_mail_DPF/) | Phone + e-mail lookup (DPF) | Selenium |
| [`Transferencias_especiais_PT/`](Transferencias_especiais_PT/) | Special transfers — action plans extraction | Selenium |
| [`Transferencias_especiais_PT_hist/`](Transferencias_especiais_PT_hist/) | Special transfers — historical plan detail | Selenium |

---

## 🚀 Getting Started

### Prerequisites

- **Python 3.8+** (some modules use 3.10+ typing)
- **Google Chrome** (most scrapers attach to / drive a live Chrome)
- **Windows** (PowerShell + Excel expected for some `.xlsx`/Outlook features)
- **Internet access** to the target government portals
### Install

```bash
git clone https://github.com/SNEAELIS/Felipe-.git
cd Felipe-
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### Run a module

Each folder is self-contained. For example:

```bash
cd CGAC_Contratos
python CGAC_Contratos.py
```

Always check the folder's own `README.md` for any module-specific input `.xlsx` format and Chrome requirements.



---

## 🛠️ Common Dependencies

| Dependency | Used for |
|---|---|
| `selenium`, `webdriver-manager` | Driving Chrome for Transferegov/SEI scrapers |
| `playwright` | Modern async browser automation (SEI dashboards, Aba Dados) |
| `pandas`, `openpyxl`, `numpy` | Excel / tabular data processing |
| `pdfplumber`, `PyPDF2`, `PyMuPDF` | PDF extraction & merge |
| `pytesseract` + `Pillow` | OCR / image handling |
| `pyautogui`, `pywin32` | Desktop automation & Outlook e-mail |
| `zeep` | SEI SOAP Web Services |
| `streamlit` | Internal SEI dashboards |

---

## 🔐 Security & Data Protection

- **No credentials are stored in this repository.** SEI/portal credentials must be provided via environment variables or local (gitignored) files.
- `*.json`, `*.xlsx`, `*.txt`, `*.log`, `.env`, `*_credentials*` are **gitignored** and must never be committed.
- All data processed is **public/official-use information** of the Brazilian federal administration-and must be handled in accordance with **LGPD** and institutional policy.
- **Never push** `.xlsx` output files, logs, `__pycache__`, or screenshots that may contain sensitive process data.



---

## 🤝 Contributing

1. Fork the repository..
2. Create a feature branch (`git checkout -b feature/my-module`).
3. Commit your changes with a clear message..
4. Open a pull request..

Please follow the existing structure — **one self-contained module per folder**, with its own `README.md` and a runnable `*.py` entry point..



---

## 📄 License

Distributed under the **MIT License**. See [LICENSE](LICENSE) for details..



---

## 👤 About

Maintained by the **SNEAELIS** team.. These automations are internal tools to support federal service-contract workflows — they are not official products of the federal government..