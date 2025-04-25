import streamlit as st
import pandas as pd
import numpy as np
import re

st.set_page_config(layout="wide")

def process_data(uploaded_file):
    if not uploaded_file:
        return None, None

    # Read file
    df = pd.read_excel(uploaded_file)

    # Drop unnecessary rows and columns
    df_cleaned = df.drop(index=0)
    df_cleaned = df_cleaned.loc[:, ~df_cleaned.columns.str.contains('^Unnamed')]

    # Rename first column
    df_cleaned = df_cleaned.rename(columns={df_cleaned.columns[0]: "Student Name"})

    # Extract categories
    categories = []
    for col in df_cleaned.columns[1:]:
        match = re.match(r"^(.*?) Q\d+", col)
        if match:
            cat = match.group(1).strip()
            if cat not in categories:
                categories.append(cat)

    # Dictionary of weights
    weights = {
        "TO BE_SER": 0.25,
        "TO LEARN": 0.20,
        "TO DO": 0.20,
        "TO DECIDE": 0.30,
        "AUTO EVAL": 0.05
    }

    # Normalize weights dictionary for matching
    normalized_weights = {k.strip().upper(): v for k, v in weights.items()}

    # Calculate raw averages and weighted scores
    for cat in categories:
        pattern = f"^{re.escape(cat)} Q\\d+"
        cols = df_cleaned.filter(regex=pattern).columns
        if not cols.empty:
            raw_avg = df_cleaned[cols].astype(float).mean(axis=1)
            df_cleaned[f"{cat} Average"] = raw_avg

            # Normalize the category name
            cat_key = cat.strip().upper()
            wt = normalized_weights.get(cat_key, None)
            weighted = raw_avg * wt if wt is not None else raw_avg
            avg_col = f"Average {cat}"
            df_cleaned[avg_col] = weighted

    # Calculate final average
    weight_cols = [col for col in df_cleaned.columns if col.startswith("Average ")]
    df_cleaned["Final Average"] = df_cleaned[weight_cols].sum(axis=1)

    # Prepare summary report
    report_cols = ["Student Name"] + weight_cols + ["Final Average"]
    df_report = df_cleaned[report_cols]

    return df_cleaned, df_report

# Streamlit interface
st.title("Rubric Evaluator with Weighting System")

uploaded_file = st.file_uploader("Upload rubric Excel file", type=["xlsx"])

if uploaded_file:
    df_cleaned, df_report = process_data(uploaded_file)

    if df_cleaned is not None:
        st.subheader("Full Data with Raw and Weighted Averages")
        st.dataframe(df_cleaned, use_container_width=True)

        st.subheader("Summary Report (Weighted Averages)")
        st.dataframe(df_report, use_container_width=True)

        # Download button
        st.download_button(
            label="Download Report as Excel",
            data=df_report.to_excel(index=False, engine='openpyxl'),
            file_name="summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
