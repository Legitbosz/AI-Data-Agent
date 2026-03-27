import pandas as pd
import sys
sys.path.insert(0, '.')

from exporters import fix_chart_values, resolve_all_charts

# Simulate exact AI response that returned North=107,000
analysis = {
    "chart_title": "Sales Analysis",
    "charts_data": [
        {
            "title": "Sales by Region",
            "categories": ["West", "North", "South", "East"],
            "series": {"Sales": [163000, 107000, 79000, 74000]},
            "label_format": "currency",
            "chart_type": "bar",
            "x_title": "Region",
            "y_title": "Sales ($)"
        }
    ]
}

# Load your actual CSV — adjust path if needed
try:
    df = pd.read_csv("sample_data.csv")
except:
    try:
        df = pd.read_excel("sample_data.xlsx")
    except:
        # Hardcode the dataset from screenshots
        data = {
            'Month':   ['Jan','Jan','Jan','Feb','Feb','Feb','Mar','Mar','Mar','Apr','Apr','Apr'],
            'Product': ['Laptop','Phone','Tablet','Laptop','Phone','Tablet','Laptop','Phone','Tablet','Laptop','Phone','Tablet'],
            'Sales':   [45000,32000,18000,52000,29000,21000,61000,38000,15000,48000,41000,23000],
            'Units':   [90,160,72,104,145,84,122,190,60,96,205,92],
            'Region':  ['North','South','East','North','South','East','North','West','East','North','South','West']
        }
        df = pd.DataFrame(data)

print("Dataframe columns:", list(df.columns))
print("Dataframe shape:", df.shape)
print("\nActual Region totals:")
print(df.groupby('Region')['Sales'].sum().sort_values(ascending=False))

print("\nresolve_all_charts result:")
result = resolve_all_charts(analysis, df)
for cd in result:
    print(f"\n{cd['title']}")
    for cat, val in zip(cd['categories'], list(cd['series'].values())[0]):
        print(f"  {cat}: {val:,.0f}")
