import streamlit as st
import pandas as pd
import os
import shutil
from agent import analyze_data, analyze_two_datasets, run_analysis_code
from exporters import export_to_excel, export_to_word, export_to_pptx
from cleaner import clean_dataframe, get_data_quality_report

st.set_page_config(page_title="AI Data Agent", layout="wide")

st.markdown("""
<style>
div[data-baseweb="select"] * { cursor: default !important; }
div[data-baseweb="select"] [role="option"] { cursor: default !important; }
div[data-baseweb="select"] [role="listbox"] { cursor: default !important; }
div[data-baseweb="select"] input { cursor: default !important; }
</style>
""", unsafe_allow_html=True)

st.title("AI Data Analysis Agent")
st.caption("Upload files, clean data, ask questions, export to Office")

for key, default in {
    "messages": [], "df1": None, "df2": None,
    "name1": "Dataset 1", "name2": "Dataset 2",
    "last_analysis": None
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def load_file(uploaded, header_row=0):
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded, header=header_row)
    else:
        df = pd.read_excel(uploaded, header=header_row)
    # Clean column names — strip whitespace and non-breaking spaces
    df.columns = df.columns.astype(str).str.replace(r'[ \s]+', ' ', regex=True).str.strip()
    return df


def preview_raw(uploaded, n=10):
    try:
        uploaded.seek(0)
        if uploaded.name.endswith(".csv"):
            return pd.read_csv(uploaded, header=None, nrows=n)
        else:
            return pd.read_excel(uploaded, header=None, nrows=n)
    except Exception:
        return None
    finally:
        uploaded.seek(0)


# ── MAIN AREA: uploaders run first so session_state is populated ───────────────
st.subheader("Upload Files")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**File 1**")
    file1 = st.file_uploader("Upload first file", type=["csv","xlsx","xls"], key="file1")
    if file1:
        raw1 = preview_raw(file1)
        st.session_state._raw1      = raw1
        st.session_state._file1_obj = file1
        if raw1 is not None:
            st.markdown("**Preview (raw):**")
            st.dataframe(raw1, use_container_width=True)
        file1.seek(0)
        header1 = int(st.session_state.get("header1_sb", 0) or 0)
        st.session_state.df1   = load_file(file1, header1)
        st.session_state.name1 = file1.name.rsplit(".", 1)[0]
        st.success(f"Loaded: {st.session_state.df1.shape[0]} rows x {st.session_state.df1.shape[1]} cols")
        st.dataframe(st.session_state.df1.head(5), use_container_width=True)

with col2:
    st.markdown("**File 2 (optional — for comparison)**")
    file2 = st.file_uploader("Upload second file", type=["csv","xlsx","xls"], key="file2")
    if file2:
        raw2 = preview_raw(file2)
        st.session_state._raw2      = raw2
        st.session_state._file2_obj = file2
        if raw2 is not None:
            st.markdown("**Preview (raw):**")
            st.dataframe(raw2, use_container_width=True)
        file2.seek(0)
        header2 = int(st.session_state.get("header2_sb", 0) or 0)
        st.session_state.df2   = load_file(file2, header2)
        st.session_state.name2 = file2.name.rsplit(".", 1)[0]
        st.success(f"Loaded: {st.session_state.df2.shape[0]} rows x {st.session_state.df2.shape[1]} cols")
        st.dataframe(st.session_state.df2.head(5), use_container_width=True)

if st.session_state.df1 is not None and st.session_state.df2 is not None:
    st.info(f"Comparison mode: **{st.session_state.name1}** vs **{st.session_state.name2}**")
elif st.session_state.df1 is not None:
    st.info(f"Single file mode: **{st.session_state.name1}**")

# ── SIDEBAR: rendered after uploaders so df is already set in session_state ────
with st.sidebar:
    if st.session_state.df1 is not None:
        with st.expander("📋 File 1 — Header Row", expanded=True):
            if "_raw1" in st.session_state and st.session_state._raw1 is not None:
                header1_sb = st.selectbox(
                    "Which row is the header?",
                    options=list(range(10)),
                    format_func=lambda x: f"Row {x} — {list(st.session_state._raw1.iloc[x].astype(str))}",
                    key="header1_sb"
                )
                if st.button("Apply Header", key="apply_header1"):
                    st.session_state._file1_obj.seek(0)
                    st.session_state.df1 = load_file(st.session_state._file1_obj, header1_sb)
                    st.rerun()

    if st.session_state.df2 is not None:
        with st.expander("📋 File 2 — Header Row", expanded=False):
            if "_raw2" in st.session_state and st.session_state._raw2 is not None:
                header2_sb = st.selectbox(
                    "Which row is the header?",
                    options=list(range(10)),
                    format_func=lambda x: f"Row {x} — {list(st.session_state._raw2.iloc[x].astype(str))}",
                    key="header2_sb"
                )
                if st.button("Apply Header", key="apply_header2"):
                    st.session_state._file2_obj.seek(0)
                    st.session_state.df2 = load_file(st.session_state._file2_obj, header2_sb)
                    st.rerun()

    if st.session_state.df1 is not None:
        with st.expander("🧹 Data Cleaning", expanded=False):
            qr = get_data_quality_report(st.session_state.df1)
            st.markdown(f"**Data Quality — {st.session_state.name1}**")
            m1, m2 = st.columns(2)
            m1.metric("Rows",       qr["total_rows"])
            m2.metric("Columns",    qr["total_columns"])
            m3, m4 = st.columns(2)
            m3.metric("Duplicates", qr["duplicate_rows"])
            m4.metric("Missing",    qr["missing_values"])
            st.dataframe(pd.DataFrame(qr["columns"]), use_container_width=True)
            st.markdown("**Select cleaning operations:**")
            options = st.multiselect(
                "Choose what to fix",
                options=[
                    "Remove duplicate rows",
                    "Drop rows with any missing values",
                    "Fill missing numbers with 0",
                    "Fill missing text with Unknown",
                    "Strip whitespace from text columns",
                    "Convert text numbers to numeric",
                    "Remove completely empty columns",
                    "Standardize column names"
                ],
                key="clean_options"
            )
            if st.button("Apply Cleaning", key="apply_cleaning") and options:
                cleaned_df, report = clean_dataframe(st.session_state.df1.copy(), options)
                st.session_state.df1 = cleaned_df
                st.success("Cleaning complete!")
                for line in report:
                    st.markdown(f"- {line}")
                st.dataframe(st.session_state.df1.head(10), use_container_width=True)
                csv = st.session_state.df1.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Download Cleaned CSV",
                    csv,
                    file_name=f"{st.session_state.name1}_cleaned.csv",
                    mime="text/csv",
                    key="download_cleaned"
                )

st.divider()

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        if msg["role"] == "assistant":
            st.write(msg["answer"])
            if msg.get("findings"):
                st.subheader("Key Findings")
                for f in msg["findings"]:
                    st.markdown(f"- {f}")
            if msg.get("chart") and os.path.exists(msg["chart"]):
                st.image(msg["chart"], width=800)
        else:
            st.write(msg["content"])

question = st.chat_input("Ask a question about your data...", key="main_chat")

if question:
    if st.session_state.df1 is None:
        st.warning("Please upload at least one file first.")
    else:
        st.session_state.messages.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.write(question)
        with st.chat_message("assistant"):
            with st.spinner("Analyzing..."):
                comparing = st.session_state.df2 is not None
                if comparing:
                    analysis = analyze_two_datasets(
                        st.session_state.df1, st.session_state.name1,
                        st.session_state.df2, st.session_state.name2,
                        question, st.session_state.messages
                    )
                    run_analysis_code(
                        analysis.get("python_code", ""),
                        df1=st.session_state.df1,
                        df2=st.session_state.df2
                    )
                else:
                    analysis = analyze_data(
                        st.session_state.df1, question, st.session_state.messages
                    )
                    run_analysis_code(
                        analysis.get("python_code", ""),
                        df=st.session_state.df1
                    )
                st.session_state.last_analysis = analysis

            st.write(analysis.get("answer", ""))
            findings = analysis.get("key_findings", [])
            if findings:
                st.subheader("Key Findings")
                for f in findings:
                    st.markdown(f"- {f}")

            chart_path = f"chart_{len(st.session_state.messages)}.png"
            if os.path.exists("chart.png"):
                shutil.copy("chart.png", chart_path)
                st.image(chart_path, width=800)
            else:
                chart_path = None

            st.session_state.messages.append({
                "role": "assistant",
                "answer": analysis.get("answer", ""),
                "findings": findings,
                "chart": chart_path
            })

        if st.session_state.last_analysis:
            st.divider()
            st.subheader("Export Last Analysis")
            c1, c2, c3 = st.columns(3)
            with c1:
                ep = export_to_excel(
                    st.session_state.df1, st.session_state.last_analysis, "chart.png",
                    df2=st.session_state.df2,
                    name1=st.session_state.name1, name2=st.session_state.name2
                )
                with open(ep, "rb") as f:
                    st.download_button("Download Excel (.xlsx)", f, file_name="report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_excel")
            with c2:
                wp = export_to_word(
                    st.session_state.df1, st.session_state.last_analysis, "chart.png",
                    df2=st.session_state.df2,
                    name1=st.session_state.name1, name2=st.session_state.name2
                )
                with open(wp, "rb") as f:
                    st.download_button("Download Word (.docx)", f, file_name="report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="dl_word")
            with c3:
                pp = export_to_pptx(
                    st.session_state.last_analysis, "chart.png",
                    df=st.session_state.df1, df2=st.session_state.df2,
                    name1=st.session_state.name1, name2=st.session_state.name2
                )
                with open(pp, "rb") as f:
                    st.download_button("Download PowerPoint (.pptx)", f, file_name="report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="dl_pptx")

if st.session_state.messages:
    if st.button("Clear Chat", key="clear_chat"):
        st.session_state.messages = []
        st.session_state.last_analysis = None
        st.rerun()
