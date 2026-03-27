<div align="center">

# 🤖 AI Data Analysis Agent

### Turn any dataset into actionable insights — instantly.

**Powered by Claude AI (Anthropic) · Built with Streamlit · Exports to Excel, Word & PowerPoint**

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.x-red.svg)](https://streamlit.io)
[![Claude](https://img.shields.io/badge/Claude-Sonnet--4-purple.svg)](https://anthropic.com)
[![License](https://img.shields.io/badge/License-Private-gray.svg)](#)

</div>

---

## 📌 What Is This?

The **AI Data Analysis Agent** is a full-stack intelligent data analytics platform. Upload any CSV or Excel file, ask a question in plain English, and the agent:

- **Understands your data** — column types, structure, currency, scale
- **Generates accurate charts** — bar, line, pie, multi-subplot
- **Computes verified insights** — key findings backed by actual pandas calculations, not AI guesses
- **Exports professional reports** — fully editable Excel workbooks, Word documents, and PowerPoint presentations

No SQL. No formulas. No manual chart building. Just ask.

---

## 🎯 Core Capabilities

### 🗣️ Natural Language Analysis
Ask anything about your data:
- *"Show me total revenue by booking center"*
- *"Which agency generates the most sales? Give me a breakdown"*
- *"Compare monthly revenue trends and show distribution by ad category"*
- *"What is the average bill amount by discount range?"*

The agent interprets the question, writes Python code, executes it, and returns verified results.

### 📊 Chart Generation
| Chart Type | Description |
|------------|-------------|
| Bar / Column | Category comparisons, rankings, totals |
| Line | Trends over time, monthly/quarterly patterns |
| Pie | Distribution, percentage breakdowns |
| Multi-subplot | Up to 4 charts in one analysis |

All charts are rendered using matplotlib on the platform and exported as **native editable charts** in Office formats — not images.

### 🔢 Verified Data Accuracy
A critical design principle of this agent:

> **AI text is never trusted for numbers. All values are recomputed from the actual dataset.**

- `findings_code` — AI writes Python code that is *executed* to generate key findings
- `fix_chart_values()` — recomputes all chart data directly from the dataframe using pandas groupby
- `best_numeric_col()` — intelligently selects the correct financial column (Bill Amt, Revenue, Sales, etc.) even when column names don't match the series label
- Verified Data Summary in exports always reflects true dataset values

### 🌍 Multi-Currency Detection
Automatically detects and applies the correct currency symbol:
- ₦ Nigerian Naira
- $ US Dollar
- £ British Pound
- € Euro
- Scans column names, answer text, and key findings for currency context

### 🧹 Data Cleaning Suite
Before analysis, clean your data with one click:

| Operation | Description |
|-----------|-------------|
| Remove duplicate rows | Deduplicates entire dataset |
| Drop rows with missing values | Removes incomplete records |
| Fill missing numbers with 0 | Handles null numeric fields |
| Fill missing text with "Unknown" | Handles null categorical fields |
| Strip whitespace | Cleans messy text columns |
| Convert text numbers to numeric | Parses formatted number strings |
| Remove empty columns | Drops fully null columns |
| Standardize column names | Lowercase, underscore format |

### 📤 Export Engine
Every export contains:

**Excel (.xlsx)**
- `Summary` sheet — Verified Data Summary (always accurate) + AI Answer + Key Findings
- `Chart 1`, `Chart 2`... — Native editable chart + linked data table on same sheet
- `Raw Data` — Full original dataset

**Word (.docx)**
- Full answer and key findings
- Editable data tables with chart instructions
- Data sample (first 20 rows)

**PowerPoint (.pptx)**
- Title slide
- Analysis Answer slide
- Key Findings slide
- One native editable chart slide per subplot

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        app.py (UI)                          │
│   Streamlit interface — upload, clean, chat, export         │
└──────────────┬──────────────────────────┬───────────────────┘
               │                          │
     ┌─────────▼──────────┐    ┌──────────▼──────────┐
     │     agent.py        │    │    exporters.py      │
     │                     │    │                      │
     │  • analyze_data()   │    │  • export_to_excel() │
     │  • analyze_two_     │    │  • export_to_word()  │
     │    datasets()       │    │  • export_to_pptx()  │
     │  • run_analysis_    │    │  • build_chart_xml() │
     │    code()           │    │  • fix_chart_values()│
     │  • compute_key_     │    │  • resolve_all_      │
     │    findings()       │    │    charts()          │
     │                     │    │  • detect_currency() │
     └─────────┬───────────┘    └──────────────────────┘
               │
     ┌─────────▼───────────┐    ┌──────────────────────┐
     │   Claude Sonnet API  │    │     cleaner.py        │
     │                      │    │                       │
     │  • Understands data  │    │  • clean_dataframe()  │
     │  • Generates charts  │    │  • get_data_quality_  │
     │  • Returns JSON with │    │    report()           │
     │    charts_data,      │    └──────────────────────┘
     │    findings_code,    │
     │    currency_symbol   │
     └──────────────────────┘
```

### Data Flow

```
User uploads file
       ↓
Column names normalized (strip \xa0, extra spaces)
       ↓
User asks question
       ↓
Claude API receives: dataset summary + question + history
       ↓
Claude returns JSON:
  - answer (text)
  - findings_code (executable Python)
  - python_code (chart generation)
  - charts_data (structure + labels)
  - currency_symbol
       ↓
findings_code is EXECUTED → accurate key findings
python_code is EXECUTED → chart.png generated
       ↓
On export:
  fix_chart_values() recomputes all values from dataframe
  build_chart_xml() generates correct OOXML from scratch
  inject_chart_xmls() patches the Excel file
       ↓
User downloads report
```

---

## 📁 File Structure

```
AI-Data-Agent/
│
├── app.py              # Main Streamlit application
│                       # — File upload with header row picker (sidebar)
│                       # — Data cleaning panel (sidebar)
│                       # — Chat interface
│                       # — Export buttons
│
├── agent.py            # Claude AI integration
│                       # — build_prompt_single() / build_prompt_two()
│                       # — analyze_data() / analyze_two_datasets()
│                       # — compute_key_findings() — executes findings_code
│                       # — run_analysis_code() — executes chart Python
│
├── exporters.py        # Office export engine
│                       # — detect_currency()
│                       # — FINANCIAL_KEYWORDS / DIMENSION_KEYWORDS
│                       # — best_numeric_col() — smart column selection
│                       # — fix_chart_values() — recomputes from dataframe
│                       # — resolve_all_charts() — builds chart list
│                       # — build_chart_xml() — scratch OOXML generation
│                       # — inject_chart_xmls() — patches xlsx file
│                       # — export_to_excel/word/pptx()
│
├── cleaner.py          # Data cleaning utilities
│                       # — clean_dataframe()
│                       # — get_data_quality_report()
│
├── cli.py              # Command-line interface
│
├── .env                # API key (never committed)
├── .gitignore
└── README.md
```

---

## 🚀 Installation & Setup

### Requirements
- Python 3.10 or higher
- Anthropic API key → [Get one here](https://console.anthropic.com/settings/keys)

### Step 1 — Clone

```bash
git clone https://github.com/Legitbosz/AI-Data-Agent.git
cd AI-Data-Agent
```

### Step 2 — Virtual Environment

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Mac / Linux
python3 -m venv venv
source venv/bin/activate
```

### Step 3 — Install Dependencies

```bash
pip install streamlit pandas anthropic python-dotenv openpyxl python-docx python-pptx matplotlib lxml
```

### Step 4 — Configure API Key

Create a `.env` file in the project root:

```env
ANTHROPIC_API_KEY=sk-ant-your-key-here
```

### Step 5 — Run

```bash
streamlit run app.py
```

Open your browser at `http://localhost:8501`

---

## 💻 Usage Guide

### 1. Upload Your Data
- Click **Upload first file** in the sidebar
- Preview the raw data and select the correct header row
- The data loads automatically with cleaned column names

### 2. Clean Your Data (Optional)
- Expand **🧹 Data Cleaning** in the sidebar
- Review the quality report (duplicates, missing values)
- Select cleaning operations and click **Apply Cleaning**
- Download the cleaned CSV if needed

### 3. Ask a Question
- Type your question in the chat input at the bottom
- Examples:
  - *"What is the total revenue by agency?"*
  - *"Show me monthly trends and top 5 clients"*
  - *"Compare bookings across centers with a pie chart"*

### 4. Export Your Report
After analysis, click:
- **Download Excel (.xlsx)** — Full workbook with editable charts
- **Download Word (.docx)** — Document report
- **Download PowerPoint (.pptx)** — Slide deck

### 5. Edit Charts in Excel
Each chart sheet contains:
- The native chart (right side)
- Its linked data table (left side, rows 3+)

**Edit any value in the data table → chart updates automatically**

---

## 🔧 Technical Notes

### Why OOXML from Scratch?
`openpyxl` generates broken chart XML:
- Uses `numRef` instead of `strRef` for text categories → Excel treats month names as series
- Sets `catAx` position to `left` instead of `bottom` → X-axis doesn't render
- `chart.legend = None` corrupts the XML silently

This agent writes correct OOXML directly, bypassing all openpyxl chart limitations.

### Why Recompute Values?
Claude (like all LLMs) frequently hallucinates numeric aggregations. For example:
- Actual: North = ₦206,000,000 | AI returned: North = ₦107,000
- Actual: West = ₦61,000 | AI returned: West = ₦163,000

`fix_chart_values()` always recomputes from the real dataframe using pandas `groupby().sum()`, completely ignoring AI-provided values.

### Why Execute findings_code?
Same reason — AI text findings contain wrong numbers. Instead of trusting the text, the agent asks Claude to write Python code that computes the findings, then *executes that code* against the real dataframe. The output is always mathematically correct.

---

## 🗺️ Roadmap

- [ ] Multi-sheet Excel file support
- [ ] SQL database connection
- [ ] Scheduled reports (email delivery)
- [ ] Dashboard mode (multiple charts on one view)
- [ ] User authentication and report history
- [ ] Custom branding / white-label exports
- [ ] API endpoint for programmatic access

---

## 🛠️ Built With

| Library | Purpose |
|---------|---------|
| [Streamlit](https://streamlit.io) | Web UI framework |
| [Anthropic Claude](https://anthropic.com) | AI analysis (claude-sonnet-4-20250514) |
| [pandas](https://pandas.pydata.org) | Data processing & computation |
| [matplotlib](https://matplotlib.org) | Chart image generation |
| [openpyxl](https://openpyxl.readthedocs.io) | Excel file handling |
| [python-docx](https://python-docx.readthedocs.io) | Word document generation |
| [python-pptx](https://python-pptx.readthedocs.io) | PowerPoint generation |
| [lxml](https://lxml.de) | XML processing |

---

## 👤 Author

**Jamil** ([@Legitbosz](https://github.com/Legitbosz))
Web3 Developer · AI Builder · Content Creator

---

## 📄 License

This project is proprietary. All rights reserved © 2025.
