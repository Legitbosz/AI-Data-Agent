import pandas as pd
from exporters import export_to_excel

df = pd.read_csv('sample_data.csv')

analysis = {
    "chart_title": "Sales by Product",
    "answer": "Test answer",
    "key_findings": ["Laptop had highest sales", "Phone had 281 units", "Tablet was lowest"]
}

path = export_to_excel(df, analysis, "chart.png", "test_report.xlsx")
print("Excel saved to:", path)