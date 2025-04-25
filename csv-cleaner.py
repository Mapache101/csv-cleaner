import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime
import math

# ——— Exact weights you use in Schoology ———
weights = {
    "AUTO EVAL":         0.05,
    "TO BE_SER":         0.05,
    "TO DECIDE_DECIDIR": 0.05,
    "TO DO_HACER":       0.40,
    "TO KNOW_SABER":     0.45
}

def custom_round(value):
    """Round .5 and above up."""
    return math.floor(value + 0.5)

def process_data(df, teacher, subject, course, level):
    # 1) Drop everything except raw assignments + Category Score columns
    to_drop = [
        "Nombre de usuario", "Username", "Promedio General",
        # we no longer drop Category Score columns here
        "Unique User ID", "Overall", "2025", "Term1 - 2025",
        "Term2- 2025", "Term3 - 2025"
    ]
    df = df.drop(columns=to_drop, errors='ignore')
    df.replace("Missing", pd.NA, inplace=True)

    # 2) Identify raw-assignment columns vs Category Score columns
    assignment_cols = []
    category_score_cols = {}
    exclusion = ["(Count in Grade)", "Ungraded"]

    for col in df.columns:
        col_str = str(col)
        if any(ph in col_str for ph in exclusion):
            continue

        # Category Score: e.g. "Term1 - 2025 - TO DO_HACER - Category Score"
        m = re.match(r".*-\s*([^–-]+?)\s*-\s*Category Score$", col_str)
        if m:
            cat = m.group(1).strip()
            category_score_cols[cat] = col_str
        elif "Grading Category:" in col_str:
            # raw assignment column
            # rename it to "<base> <category>"
            m2 = re.search(r'Grading Category:\s*([^,)]+)', col_str)
            cat2 = m2.group(1).strip() if m2 else ""
            base = col_str.split('(')[0].strip()
            new_name = f"{base} {cat2}"
            df.rename({col_str: new_name}, axis=1, inplace=True)
            assignment_cols.append(new_name)

    # 3) Pull names first
    name_terms = ["name", "first", "last"]
    name_cols = [c for c in df.columns
                 if any(t in c.lower() for t in name_terms)]

    # 4) Build our "Average <category>" from the Category Score columns
    avg_cols = []
    for cat, w in weights.items():
        raw_col = category_score_cols.get(cat)
        if not raw_col:
            continue
        avg_name = f"Average {cat}"
        # convert percent → numeric
        df[avg_name] = pd.to_numeric(df[raw_col], errors='coerce').fillna(0)
        avg_cols.append((cat, avg_name))

    # 5) Compute Final Grade = sum( avg% * weight ) then round half-up
    df["Final Grade"] = df.apply(
        lambda row: custom_round(
            sum(row[avg]*weights[cat] for cat, avg in avg_cols)
        ),
        axis=1
    )

    # 6) Final column ordering:
    ordered = []
    ordered.extend(name_cols)
    # then all raw assignments, in the order they were in the CSV
    ordered.extend([c for c in assignment_cols])
    # then each Average <cat> in weights order
    for cat, avg in avg_cols:
        ordered.append(avg)
    ordered.append("Final Grade")

    df_final = df[ordered]

    # 7) Export to Excel (identical to your existing code)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter',
                        engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:

        df_final.to_excel(writer, 'Sheet1', startrow=6, index=False)
        wb = writer.book
        ws = writer.sheets['Sheet1']

        # Header block
        b_fmt = wb.add_format({'border': 1})
        ws.write('A1', "Teacher:", b_fmt); ws.write('B1', teacher, b_fmt)
        ws.write('A2', "Subject:", b_fmt); ws.write('B2', subject, b_fmt)
        ws.write('A3', "Class:",   b_fmt); ws.write('B3', course,  b_fmt)
        ws.write('A4', "Level:",   b_fmt); ws.write('B4', level,   b_fmt)
        ws.write('A5', datetime.now().strftime("%y-%m-%d"), b_fmt)

        # Formats
        header_fmt = wb.add_format({
            'bold': True, 'border': 1,
            'rotation': 90, 'shrink': True, 'text_wrap': True
        })
        avg_hdr = wb.add_format({
            'bold': True, 'border': 1,
            'rotation': 90, 'shrink': True,
            'text_wrap': True, 'bg_color': '#ADD8E6'
        })
        avg_data = wb.add_format({
            'border': 1, 'bg_color': '#ADD8E6', 'num_format': '0'
        })
        final_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#90EE90'})

        # Write column headers
        for idx, col in enumerate(df_final.columns):
            if col == "Final Grade":
                fmt = final_fmt
            elif col.startswith("Average "):
                fmt = avg_hdr
            else:
                fmt = header_fmt
            ws.write(6, idx, col, fmt)

        # Write data rows
        for c_idx, col in enumerate(df_final.columns):
            if col == "Final Grade":
                fmt = final_fmt
            elif col.startswith("Average "):
                fmt = avg_data
            else:
                fmt = b_fmt
            for row in range(len(df_final)):
                val = df_final.iloc[row, c_idx]
                ws.write(7 + row, c_idx,
                         "" if pd.isna(val) else val,
                         fmt)

        # Adjust column widths
        for idx, col in enumerate(df_final.columns):
            low = col.lower()
            if any(t in low for t in name_terms):
                ws.set_column(idx, idx, 25)
            elif col.startswith("Average "):
                ws.set_column(idx, idx, 7)
            elif col == "Final Grade":
                ws.set_column(idx, idx, 12)
            else:
                ws.set_column(idx, idx, 5)

    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Gradebook Organizer")
    st.sidebar.markdown("""  
        1. **Ensure Schoology is set to English**  
        2. Navigate to the **course** you want to export  
        3. Click on **Gradebook**  
        4. Click the **three dots** and select **Export**  
        5. Choose **Gradebook as CSV**  
        6. **Upload** that CSV here  
        7. Fill in the fields  
        8. Click **Download Organized Gradebook (Excel)**
    """)
    st.title("Griffin CSV to Excel v2")

    # ← these must be defined before calling process_data()
    teacher = st.text_input("Enter teacher's name:")
    subject = st.text_input("Enter subject area:")
    course  = st.text_input("Enter class:")
    level   = st.text_input("Enter level:")

    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            df.columns = df.columns.astype(str)
            out = process_data(df, teacher, subject, course, level)
            st.download_button(
                "Download Organized Gradebook (Excel)",
                data=out,
                file_name="final_cleaned_gradebook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Processing completed!")
        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
