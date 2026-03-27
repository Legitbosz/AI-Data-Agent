import pandas as pd
import numpy as np


def clean_dataframe(df, options):
    """
    Clean a dataframe based on selected options.
    options is a list of strings like:
    ['Remove duplicate rows', 'Fill missing numbers with 0', ...]
    """
    report = []
    original_rows = len(df)

    if "Remove duplicate rows" in options:
        before = len(df)
        df = df.drop_duplicates()
        removed = before - len(df)
        report.append(f"Removed {removed} duplicate rows")

    if "Drop rows with any missing values" in options:
        before = len(df)
        df = df.dropna()
        removed = before - len(df)
        report.append(f"Dropped {removed} rows with missing values")

    if "Fill missing numbers with 0" in options:
        numeric_cols = df.select_dtypes(include="number").columns
        filled = df[numeric_cols].isna().sum().sum()
        df[numeric_cols] = df[numeric_cols].fillna(0)
        report.append(f"Filled {filled} missing numeric values with 0")

    if "Fill missing text with Unknown" in options:
        text_cols = df.select_dtypes(exclude="number").columns
        filled = df[text_cols].isna().sum().sum()
        df[text_cols] = df[text_cols].fillna("Unknown")
        report.append(f"Filled {filled} missing text values with 'Unknown'")

    if "Strip whitespace from text columns" in options:
        text_cols = df.select_dtypes(include="object").columns
        for col in text_cols:
            df[col] = df[col].astype(str).str.strip()
        report.append(f"Stripped whitespace from {len(text_cols)} text columns")

    if "Convert text numbers to numeric" in options:
        converted = 0
        for col in df.columns:
            if df[col].dtype == object:
                try:
                    converted_col = pd.to_numeric(
                        df[col].astype(str).str.replace(",", "").str.replace("$", "").str.strip(),
                        errors="coerce"
                    )
                    if converted_col.notna().sum() > len(df) * 0.5:
                        df[col] = converted_col
                        converted += 1
                except Exception:
                    pass
        report.append(f"Converted {converted} columns to numeric")

    if "Remove completely empty columns" in options:
        before = len(df.columns)
        df = df.dropna(axis=1, how="all")
        removed = before - len(df.columns)
        report.append(f"Removed {removed} empty columns")

    if "Standardize column names" in options:
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
            .str.replace(r"[^a-z0-9]+", "_", regex=True)
            .str.strip("_")
        )
        report.append("Standardized all column names to lowercase with underscores")

    final_rows = len(df)
    report.append(f"Final dataset: {final_rows} rows x {len(df.columns)} columns")

    return df, report


def get_data_quality_report(df):
    """Return a quick quality summary of the dataframe."""
    report = {}
    report["total_rows"]    = len(df)
    report["total_columns"] = len(df.columns)
    report["duplicate_rows"]= int(df.duplicated().sum())
    report["missing_values"]= int(df.isna().sum().sum())

    col_info = []
    for col in df.columns:
        missing = int(df[col].isna().sum())
        dtype   = str(df[col].dtype)
        unique  = int(df[col].nunique())
        col_info.append({
            "column":  col,
            "type":    dtype,
            "missing": missing,
            "unique":  unique
        })
    report["columns"] = col_info
    return report