import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime
import math  # for our half-up rounding

# Weights per category as defined by the Bolivian law
weights = {
    "Auto eval":      0.05,
    "TO BE_SER":      0.05,
    "TO DECIDE_DECIDIR": 0.05,
    "TO DO_HACER":    0.40,
    "TO KNOW_SABER":  0.45
}

def custom_round(value):
    """Round .5 and above up, everything else down."""
    return math.floor(value + 0.5)

def process_data(df, teacher, subject, course, level):
    # --- 1) Drop columns we don’t need ---
    to_drop = [
        "Nombre de usuario", "Username", "Promedio General",
        "Term1 - 2024", "Term1 - 2024 - AUTO EVAL TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO DECIDE_DECIDIR - Puntuación de categoría",
        "Term1 - 2024 - TO DO_HACER - Puntuación de categoría",
        "Term1 - 2024 - TO KNOW_SABER - Puntuación de categoría",
        "Unique User ID", "Overall", "2025", "Term1 - 2025",
        "Term2- 2025", "Term3 - 2025"
    ]
    df = df.drop(columns=to_drop, errors='ignore')

    # Treat “Missing” as a zero
    df.replace("Missing", pd.NA, inplace=True)

    exclusion_phrases = ["(Count in Grade)", "Category Score", "Ungraded"]
    columns_info = []
    general_columns = []
    cols_to_remove = {"ID de usuario único", "ID de usuario unico"}

    # --- 2) Identify each assignment column + its category + its max points ---
    for i, col in enumerate(df.columns):
        col = str(col)
        if col in cols_to_remove or any(ph in col for ph in exclusion_phrases):
            continue

        if "Grading Category:" in col:
            # extract category name
            m_cat = re.search(r'Grading Category:\s*([^,)]+)', col)
            category = m_cat.group(1).strip() if m_cat else "Unknown"
            # extract max points
            m_pts = re.search(r'Max Points:\s*([\d\.]+)', col)
            max_pts = float(m_pts.group(1)) if m_pts else 0.0

            base_name = col.split('(')[0].strip()
            new_name = f"{base_name} {category}".strip()

            columns_info.append({
                'original': col,
                'new_name': new_name,
                'category': category,
                'seq_num': i,
                'max_points': max_pts
            })
        else:
            general_columns.append(col)

    # --- 3) Reorder so “name” comes first, then everything else, then all assignment cols ---
    name_terms = ["name", "first", "last"]
    name_cols = [c for c in general_columns if any(t in c.lower() for t in name_terms)]
    other_cols = [c for c in general_columns if c not in name_cols]
    general_reordered = name_cols + other_cols

    sorted_assignments = sorted(columns_info, key=lambda x: x['seq_num'])
    df_clean = df[general_reordered + [d['original'] for d in sorted_assignments]].copy()
    df_clean.rename({d['original']: d['new_name'] for d in columns_info},
                    axis=1, inplace=True)

    # --- 4) Group assignments by category ---
    groups = {}
    for d in columns_info:
        groups.setdefault(d['category'], []).append(d)

    # Now for each category, compute the weighted contribution:
    final_cols = []
    for cat, items in groups.items():
        names = [d['new_name'] for d in items]
        # raw points, converting blanks to zeros
        raw = df_clean[names].apply(pd.to_numeric, errors='coerce').fillna(0)
        total_earned   = raw.sum(axis=1)
        total_possible = sum(d['max_points'] for d in items) or 1.0  # avoid div/0
        pct = total_earned / total_possible * 100
        wt  = weights.get(cat, 0)

        # this is the category’s contribution to the final percentage
        contrib = pct * wt
        col_name = f"Average {cat}"
        df_clean[col_name] = contrib
        final_cols.append(col_name)

    # --- 5) Final grade = sum of all category contributions, rounded half-up ---
    df_clean["Final Grade"] = (
        df_clean[final_cols].sum(axis=1)
                   .apply(lambda x: custom_round(x))
    )

    # --- 6) Export exactly as before into Excel ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter',
                        engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        df_clean.to_excel(writer, 'Sheet1', startrow=6, index=False)
        wb = writer.book
        ws = writer.sheets['Sheet1']

        # Header info
        b_fmt     = wb.add_format({'border': 1})
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

        # Write headers
        for idx, col in enumerate(df_clean.columns):
            fmt = header_fmt
            if col.startswith("Average "): fmt = avg_hdr
            if col == "Final Grade":      fmt = final_fmt
            ws.write(6, idx, col, fmt)

        # Write data rows
        for col_idx, col in enumerate(df_clean.columns):
            if col.startswith("Average "):
                fmt = avg_data
            elif col == "Final Grade":
                fmt = final_fmt
            else:
                fmt = b_fmt

            for row_offset in range(len(df_clean)):
                val = df_clean.iloc[row_offset, col_idx]
                ws.write(7 + row_offset, col_idx,
                         "" if pd.isna(val) else val,
                         fmt)

        # Adjust column widths
        for idx, col in enumerate(df_clean.columns):
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
        4. Click the **three dots** on the top-right corner and select **Export**  
        5. Choose **Gradebook as CSV**  
        6. **Upload** that CSV file to this program  
        7. Fill in the required fields  
        8. Click **Download Organized Gradebook (Excel)**  
        9. 🎉 **Enjoy!**
    """)
    st.title("Griffin CSV to Excel v2 ")
    teacher = st.text_input("Enter teacher's name:")
    subject = st.text_input("Enter subject area:")
    course  = st.text_input("Enter class:")
    level   = st.text_input("Enter level:")
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            df.columns = df.columns.astype(str)
            output = process_data(df, teacher, subject, course, level)
            st.download_button(
                "Download Organized Gradebook (Excel)",
                data=output,
                file_name="final_cleaned_gradebook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Processing completed!")
        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
