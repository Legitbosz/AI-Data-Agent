import pandas as pd
import sys
from agent import analyze_data, run_analysis_code
from exporters import export_to_excel, export_to_word, export_to_pptx

def main():
    print("\n🤖 AI Data Analysis Agent (CLI)\n")

    # Get file
    file_path = input("Enter path to your data file (CSV or Excel): ").strip()
    if not file_path:
        print("No file provided. Exiting.")
        sys.exit(1)

    if file_path.endswith(".csv"):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    print(f"✅ Loaded: {df.shape[0]} rows × {df.shape[1]} columns")
    print(f"   Columns: {list(df.columns)}\n")

    question = input("What do you want to know about this data? ").strip()

    print("\n⏳ Analyzing with AI...\n")
    analysis = analyze_data(df, question)

    print("=" * 60)
    print("📋 ANSWER:")
    print(analysis.get("answer", ""))
    print("\n🔑 KEY FINDINGS:")
    for f in analysis.get("key_findings", []):
        print(f"  • {f}")

    chart = run_analysis_code(analysis.get("python_code", ""), df)
    print(f"\n📊 Chart saved: {chart}")

    print("\n📥 Exporting files...")
    print("  →", export_to_excel(df, analysis, "chart.png"))
    print("  →", export_to_word(df, analysis, "chart.png"))
    print("  →", export_to_pptx(analysis, "chart.png"))

    print("\n✅ Done! Open report.xlsx / report.docx / report.pptx\n")

if __name__ == "__main__":
    main()