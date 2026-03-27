# 🤖 AI Data Analysis Agent

An intelligent data analysis platform powered by Claude AI (Anthropic). Upload any dataset, ask questions in plain English, get instant charts and insights — then export everything as fully editable Excel, Word, or PowerPoint reports.

---

## ✨ Features

### 📊 Smart Analysis
- Ask questions about your data in plain English
- AI generates accurate charts and insights automatically
- Supports single dataset and two-dataset comparison mode
- Conversation memory — ask follow-up questions in context
- Handles any dataset size and structure

### 🧹 Data Cleaning
- Duplicate row removal
- Missing value handling (fill with 0, "Unknown", or drop)
- Whitespace stripping
- Text-to-numeric conversion
- Column name standardization
- Download cleaned dataset as CSV

### 📈 Chart Types Supported
- Bar / Column charts
- Line charts
- Pie charts
- Multi-chart subplots (up to 4 per analysis)

### 📤 Export to Office Formats
- **Excel (.xlsx)** — Native editable charts linked to data tables, verified data summary, key findings, raw data sheet
- **Word (.docx)** — Full report with answer, key findings, editable data tables
- **PowerPoint (.pptx)** — Presentation-ready slides with native editable charts

### 💡 Data Accuracy
- All chart values recomputed directly from the dataset — never trusts AI estimates
- Verified Data Summary always reflects actual dataset values
- Key findings computed via executed Python code, not AI text generation
- Smart column detection for financial data (Bill Amt, Revenue, Sales, etc.)

### 🌍 Multi-Currency Support
- Auto-detects currency from data context (₦ Naira, $ USD, £ GBP, € EUR)
- Correct currency symbols in charts, labels, and summary

---

## 🚀 Getting Started

### Prerequisites
- Python 3.10+
- An [Anthropic API key](https://console.anthropic.com/settings/keys)

### Installation

```bash
# Clone the repository
git clone https://github.com/Legitbosz/AI-Data-Agent.git
cd AI-Data-Agent

# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate        # Windows
source venv/bin/activate     # Mac/Linux

# Install dependencies
pip install streamlit pandas anthropic python-dotenv openpyxl python-docx python-pptx matplotlib lxml
```

### Configuration

Create a `.env` file in the project root:

```
ANTHROPIC_API_KEY=your_api_key_here
```

### Run the App

```bash
cd C:\AI-Data-Agent
venv\Scripts\activate
streamlit run app.py
```

The app will open at `http://localhost:8501`

---

## 📁 Project Structure

```
AI-Data-Agent/
├── app.py          # Streamlit UI — main application
├── agent.py        # AI brain — Claude API integration, analysis & findings
├── exporters.py    # Export engine — Excel, Word, PowerPoint generation
├── cleaner.py      # Data cleaning module
├── cli.py          # Command-line interface (alternative to UI)
├── .env            # API key (not committed to git)
├── .gitignore
└── README.md
```

---

## 🛠️ How It Works

1. **Upload** a CSV or Excel file (up to any size)
2. **Preview** raw data and select the correct header row from the sidebar
3. **Clean** your data using the built-in cleaning tools (optional)
4. **Ask** any question about your data in plain English
5. The AI:
   - Writes and executes Python code to generate accurate charts
   - Computes key findings via executed code (not text generation)
   - Returns structured chart data for export
6. **Export** to Excel, Word, or PowerPoint with one click

---

## 📊 Supported Data Formats

| Format | Extension |
|--------|-----------|
| CSV | `.csv` |
| Excel | `.xlsx`, `.xls` |

---

## 🔧 Built With

- [Streamlit](https://streamlit.io/) — Web UI
- [Anthropic Claude](https://anthropic.com/) — AI analysis engine (claude-sonnet-4-20250514)
- [pandas](https://pandas.pydata.org/) — Data processing
- [matplotlib](https://matplotlib.org/) — Chart generation
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel export
- [python-docx](https://python-docx.readthedocs.io/) — Word export
- [python-pptx](https://python-pptx.readthedocs.io/) — PowerPoint export

---

## 📋 Example Use Cases

- **Sales Analysis** — Revenue by product, region, period
- **Financial Reporting** — Booking trends, discount analysis, revenue centers
- **Marketing Data** — Campaign performance, ad category distribution
- **Operations** — Booking patterns, executive performance, center comparisons
- **Any tabular dataset** — The AI adapts to any column structure

---

## 🔒 Security

- API keys are stored in `.env` and never committed to git
- `.gitignore` excludes `.env`, `venv/`, and `__pycache__/`

---

## 👤 Author

([@Legitbosz](https://github.com/Legitbosz))  
Web3 Developer & AI Builder

---

## 📄 License

This project is private. All rights reserved.
