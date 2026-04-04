<div align="center">

# 🤖 AI Data Analysis Agent

### Upload any dataset. Ask questions in plain English. Get answers.

**Powered by Claude AI · Built with Streamlit**

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.x-red.svg)](https://streamlit.io)
[![Claude](https://img.shields.io/badge/Claude_AI-Anthropic-purple.svg)](https://anthropic.com)

</div>

---

## What It Does

Upload a CSV, Excel, or PDF file — then ask anything about your data in plain English. The agent analyzes it, generates charts, and exports professional reports.

**Analysis** — Natural language queries with verified results and auto-generated charts (bar, line, pie, multi-subplot).

**Validation** — Automated quality checks including required fields, email/phone validation, duplicate detection, error rates, workflow tracking, and productivity metrics. Exports a full validation report in Excel.

**Export** — One-click reports in Excel (with native editable charts), Word, and PowerPoint.

**Multi-currency** — Auto-detects currency from your data (₦, $, £, €, etc.).

**PDF support** — Extracts tables from PDFs including scanned documents via OCR.

---

## Quick Start

```bash
git clone https://github.com/Legitbosz/AI-Data-Agent.git
cd AI-Data-Agent
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # Mac/Linux
pip install -r requirements.txt
streamlit run app.py
```

Create a `.env` file with your API key:

```
ANTHROPIC_API_KEY=sk-ant-...
```

---

## Requirements

- Python 3.10+
- Anthropic API key ([get one here](https://console.anthropic.com/))
- Optional: Tesseract OCR + Poppler (for scanned PDF support)

---

## Screenshots

*Coming soon*

---

## License

Private — not for redistribution.
