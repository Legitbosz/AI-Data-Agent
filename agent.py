import os
import anthropic
import pandas as pd
import json
from dotenv import load_dotenv

load_dotenv()

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))


def build_summary(df, name):
    return f"""
--- {name} ---
Shape: {df.shape[0]} rows x {df.shape[1]} columns
Columns: {list(df.columns)}
Sample (first 3 rows):
{df.head(3).to_string()}
Basic stats:
{df.describe().to_string()}
"""


def clean_json(raw):
    raw = raw.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return raw.strip()


def compute_key_findings(findings_code, df=None, df1=None, df2=None):
    """
    Execute the AI's findings_code to get accurate computed key findings.
    Returns the computed list, or falls back to AI text findings if code fails.
    """
    if not findings_code:
        return None
    local_vars = {"pd": pd}
    if df is not None:
        local_vars["df"] = df
    if df1 is not None:
        local_vars["df1"] = df1
    if df2 is not None:
        local_vars["df2"] = df2
    try:
        exec(findings_code, local_vars)
        result = local_vars.get("key_findings", None)
        if isinstance(result, list) and len(result) > 0:
            return [str(f) for f in result]
    except Exception as e:
        import logging
        logging.warning(f"[compute_key_findings] code failed: {e}")
    return None


def run_analysis_code(code, df=None, df1=None, df2=None):
    import matplotlib
    matplotlib.use('Agg')
    local_vars = {"pd": pd}
    if df is not None:
        local_vars["df"] = df
    if df1 is not None:
        local_vars["df1"] = df1
    if df2 is not None:
        local_vars["df2"] = df2
    try:
        exec(code, local_vars)
        return "chart.png"
    except Exception as e:
        return f"Error running code: {e}"


def build_prompt_single(summary, history_text, user_question):
    return f"""You are a data analyst AI. The user has uploaded a dataset and asked a question.

DATASET SUMMARY:
{summary}
{history_text}
CURRENT QUESTION:
{user_question}

Your job:
1. Answer the question clearly in plain English
   - Detect the currency from the data context (column names, values) — use ₦ for Naira, $ for USD, £ for GBP etc. Never assume $ if the data is from a non-USD context
   - Set "currency_symbol" in the JSON to the correct symbol (e.g. "₦", "$", "£") — this will be used in all chart labels and summary
2. Write Python chart code to visualize the answer
   - Use the correct currency symbol in chart labels based on the data (₦ for Nigerian data, etc.)
   - For pie charts use plt.pie() with autopct, for line charts use plt.plot(), for bar charts use plt.bar() or ax.bar()
3. Write Python findings code that computes key findings AS A LIST — this code will be executed so every number must come from actual pandas computation on the dataframe, not guessed
4. Return chart data for all subplots

For python_code (charts):
- For 1 chart: use plt.figure() and plt.savefig()
- For 2 charts: use fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
- For 3 charts: use fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(18, 6))
- For 4 charts: use fig, axes = plt.subplots(2, 2, figsize=(14, 10))
- Always end with plt.tight_layout() and plt.savefig('chart.png', dpi=150, bbox_inches='tight') and plt.close()

For findings_code — write executable Python that sets key_findings as a list of strings:
IMPORTANT: Use the exact column names from the dataset. Check the dataset summary for column names first.
Example:
key_findings = []
total_revenue = df['Bill Amt'].sum()  # use the actual column name
top_center = df.groupby('BOOKING CENTER')['Bill Amt'].sum().idxmax()
top_value = df.groupby('BOOKING CENTER')['Bill Amt'].sum().max()
key_findings.append(f"{{top_center}} leads with ₦{{top_value:,.0f}}")
key_findings.append(f"Total revenue: ₦{{total_revenue:,.0f}}")

For charts_data:
- One entry per subplot — NEVER combine subplots
- Categories = X-axis labels of THAT chart only (for pie: slice labels)
- Values = exact computed values matching categories order
- "series" key name MUST be the actual column name from the dataset (e.g. "Bill Amt", "AGREED AMOUNT") — not a generic label like "Revenue". This ensures correct data mapping on export
- "chart_type" MUST match exactly: "bar" for bar/column, "line" for line, "pie" for pie charts
- "label_format": "currency", "number", "percent", or "none"
- For pie charts, set label_format to "percent"
- Example for 2 subplots:
  [{{"title": "Sales by Product", "categories": ["Laptop","Phone","Tablet"], "series": {{"Sales": [206000,140000,77000]}}, "chart_type": "bar", "label_format": "currency", "x_title": "Product", "y_title": "Sales"}},
   {{"title": "Bookings by Center", "categories": ["ABUJA","LAGOS","KANO"], "series": {{"Bookings": [65.5,21.9,12.6]}}, "chart_type": "pie", "label_format": "percent", "x_title": "", "y_title": ""}}]

Respond in this exact JSON format:
{{
  "answer": "Your plain English answer here",
  "key_findings": ["finding 1", "finding 2", "finding 3"],
  "findings_code": "key_findings = []\\n# compute findings from df using pandas\\nkey_findings.append(...)",
  "python_code": "import pandas as pd\\nimport matplotlib\\nmatplotlib.use('Agg')\\nimport matplotlib.pyplot as plt\\n# df is available\\nplt.tight_layout()\\nplt.savefig('chart.png', dpi=150, bbox_inches='tight')\\nplt.close()",
  "chart_title": "A short chart title",
  "currency_symbol": "₦",
  "charts_data": [
    {{"title": "Chart Title", "categories": ["label1","label2"], "series": {{"Series": [val1,val2]}}, "label_format": "currency", "chart_type": "bar", "x_title": "X Label", "y_title": "Y Label"}}
  ]
}}

Only return valid JSON. No extra text outside the JSON."""


def build_prompt_two(summary1, summary2, name1, name2, history_text, user_question):
    return f"""You are a data analyst AI. The user has uploaded TWO datasets and wants to compare them.

DATASET 1:
{summary1}

DATASET 2:
{summary2}
{history_text}
CURRENT QUESTION:
{user_question}

Your job:
1. Compare the two datasets and answer the question clearly
2. Write Python chart code using df1 and df2
3. Write Python findings code that computes key findings AS A LIST using actual pandas computation on df1 and df2
4. Return chart data for all subplots

For python_code (charts):
- Both dataframes available as df1 and df2
- Use {name1!r} and {name2!r} in chart labels
- For 1 chart: use plt.figure() and plt.savefig()
- For 2 charts: use fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
- Always end with plt.tight_layout() and plt.savefig('chart.png', dpi=150, bbox_inches='tight') and plt.close()

For findings_code — write executable Python that sets key_findings as a list:
key_findings = []
# use df1 and df2 for computations
key_findings.append(f"...")

For charts_data — one entry per subplot, exact computed values.

Respond in this exact JSON format:
{{
  "answer": "Your plain English comparison answer here",
  "key_findings": ["finding 1", "finding 2", "finding 3"],
  "findings_code": "key_findings = []\\n# compute from df1, df2\\nkey_findings.append(...)",
  "python_code": "import pandas as pd\\nimport matplotlib\\nmatplotlib.use('Agg')\\nimport matplotlib.pyplot as plt\\n# df1 and df2 available\\nplt.tight_layout()\\nplt.savefig('chart.png', dpi=150, bbox_inches='tight')\\nplt.close()",
  "chart_title": "A short comparison chart title",
  "currency_symbol": "₦",
  "charts_data": [
    {{"title": "Chart Title", "categories": ["label1","label2"], "series": {{"Series": [val1,val2]}}, "label_format": "currency", "chart_type": "bar", "x_title": "X Label", "y_title": "Y Label"}}
  ]
}}

Only return valid JSON. No extra text outside the JSON."""


def analyze_data(df, user_question, history=[]):
    summary = build_summary(df, "Dataset")

    history_text = ""
    if history:
        history_text = "\n\nPREVIOUS CONVERSATION:\n"
        for msg in history[-6:]:
            if msg["role"] == "user":
                history_text += f"User asked: {msg['content']}\n"
            elif msg["role"] == "assistant":
                history_text += f"You answered: {msg.get('answer', '')}\n"

    prompt = build_prompt_single(summary, history_text, user_question)

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    result = json.loads(clean_json(message.content[0].text))

    # Execute findings_code to get accurate computed findings
    computed = compute_key_findings(result.get("findings_code", ""), df=df)
    if computed:
        result["key_findings"] = computed

    import logging
    logging.warning(f"[AGENT] charts_data: {result.get('charts_data')}")
    logging.warning(f"[AGENT] key_findings: {result.get('key_findings')}")
    return result


def analyze_two_datasets(df1, name1, df2, name2, user_question, history=[]):
    summary1 = build_summary(df1, name1)
    summary2 = build_summary(df2, name2)

    history_text = ""
    if history:
        history_text = "\n\nPREVIOUS CONVERSATION:\n"
        for msg in history[-6:]:
            if msg["role"] == "user":
                history_text += f"User asked: {msg['content']}\n"
            elif msg["role"] == "assistant":
                history_text += f"You answered: {msg.get('answer', '')}\n"

    prompt = build_prompt_two(summary1, summary2, name1, name2, history_text, user_question)

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    result = json.loads(clean_json(message.content[0].text))

    # Execute findings_code to get accurate computed findings
    computed = compute_key_findings(result.get("findings_code", ""), df1=df1, df2=df2)
    if computed:
        result["key_findings"] = computed

    import logging
    logging.warning(f"[AGENT] charts_data: {result.get('charts_data')}")
    logging.warning(f"[AGENT] key_findings: {result.get('key_findings')}")
    return result
