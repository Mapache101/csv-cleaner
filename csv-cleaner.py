import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime
import math  # for floor()

# Weights per category as defined by the Bolivian law
weights = {
    "Auto eval": 0.05,
    "TO BE_SER": 0.05,
    "TO DECIDE_DECIDIR": 0.05,
    "TO DO_HACER": 0.40,
    "TO KNOW_SABER": 0.45
}

# Custom round-half-up: only .5 and above goes up; everything else rounds down
def custom_round(value):
    return math.floor(value + 0.5)


def process_data(df, teacher, subject, course, level):
    # Drop non-assignment columns
    columns_to_drop = [
        "Nombre de usuario", "Username", "Promedio General",
        "Term1 - 2024", "Term1 - 2024 - AUTO EVAL TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO DECIDE_DECIDIR - Puntuación de categoría",
        "Term1 - 2024 - TO DO_HACER - Puntuación de categoría",
        "Term1 - 2024 - TO KNOW_SABER - Puntuación de categoría",
        "Unique User ID", "Overall", "2025", "Term1 - 2025",
        "Term2- 2025", "Term3 - 2025"
    ]
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore')

    # Treat "Missing" as blank
    df.replace("Missing", pd.NA, inplace=True)

    # Capture Schoology summary columns for each category, normalized
    summary_cols = {}
    for col in df.columns:
        if col.endswith(' - Category Score'):
            parts = col.split(' - ')
            if len(parts) >= 2:
                cat_raw = parts[-2].strip()
                # normalize to weights key (case-insensitive)
                cat = next((k for k in weights if k.lower() == cat_raw.lower()), cat_raw)
                summary_cols[cat] = col

    # Parse raw assignment columns
    exclusion_phrases = ["(Count in Grade)", "Category Score", "Ungraded"]
    columns_info = []
    general_columns = []
    cols_to_remove = {"ID de usuario único", "ID de usuario unico"}
    for i, col in enumerate(df.columns):
        if col in cols_to_remove or any(ph in col for ph in exclusion_phrases):
            continue
        if "Grading Category:" in col:
            m_cat = re.search(r'Grading Category:\s*([^,)]+)', col)
            cat_raw = m_cat.group(1).strip() if m_cat else "Unknown"
            # normalize category to weights key
            category = next((k for k in weights if k.lower() == cat_raw.lower()), cat_raw)
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

    # Identify name columns
    name_terms = ["first name", "last name", "name", "first", "last"]
    name_cols = [c for c in general_columns if any(t in c.lower() for t in name_terms)]
    other_cols = [c for c in general_columns if c not in name_cols]

    # Build cleaned DataFrame with renamed raw columns
    sorted_coded = sorted(columns_info, key=lambda x: x['seq_num'])
    raw_names = [d['new_name'] for d in sorted_coded]
    new_order = name_cols + other_cols + [d['original'] for d in sorted_coded]
    df_cleaned = df[new_order].copy()
    df_cleaned.rename({d['original']: d['new_name'] for d in columns_info}, axis=1, inplace=True)

    # Group by normalized category
    groups = {}
    for d in columns_info:
        groups.setdefault(d['category'], []).append(d)

    # Compute weighted averages using summary columns
    avg_cols = []
    for cat in weights:  # ensure consistent order if needed
        avg_col = f"Average {cat}"
        src = summary_cols.get(cat)
        if src:
            pct = pd.to_numeric(df[src], errors='coerce')
            wt = weights[cat]
            df_cleaned[avg_col] = pct * wt
        else:
            df_cleaned[avg_col] = pd.NA
        avg_cols.append(avg_col)

    # Build final column order: first/last, then each category raw+avg, final grade
    desired_cats = ["TO BE_SER", "TO DECIDE_DECIDIR", "TO DO_HACER", "TO KNOW_SABER", "Auto eval"]
    final_order = []
    final_order.extend(name_cols)
    for cat in desired_cats:
        # raw tasks for this category
        for d in sorted(groups.get(cat, []), key=lambda x: x['seq_num']):
            final_order.append(d['new_name'])
        # average column
        final_order.append(f"Average {cat}")
    final_order.append("Final Grade")

    # Compute final grade
    df_cleaned["Final Grade"] = df_cleaned.apply(
        lambda row: custom_round(sum(row[f"Average {cat}"] for cat in desired_cats if pd.notna(row[f"Average {cat}"]))),
        axis=1)

    # Reorder and return
    df_final = df_cleaned[final_order]

    # Export to Excel unchanged
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        df_final.to_excel(writer, 'Sheet1', startrow=6, index=False)
        wb = writer.book
        ws = writer.sheets['Sheet1']

        header_fmt = wb.add_format({'bold': True, 'border': 1, 'rotation': 90, 'shrink': True, 'text_wrap': True})
        avg_hdr = wb.add_format({'bold': True, 'border': 1, 'rotation': 90, 'shrink': True, 'text_wrap': True, 'bg_color': '#ADD8E6'})
        avg_data = wb.add_format({'border': 1, 'bg_color': '#ADD8E6', 'num_format': '0'})
        final_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#90EE90'})
        b_fmt = wb.add_format({'border': 1})

        ws.write('A1', "Teacher:", b_fmt); ws.write('B1', teacher, b_fmt)
        ws.write('A2', "Subject:", b_fmt); ws.write('B2', subject, b_fmt)
        ws.write('A3', "Class:", b_fmt); ws.write('B3', course, b_fmt)
        ws.write('A4', "Level:", b_fmt); ws.write('B4', level, b_fmt)
        ws.write('A5', datetime.now().strftime("%y-%m-%d"), b_fmt)

        for idx, col in enumerate(df_final.columns):
            fmt = header_fmt
            if col.startswith("Average "):
                fmt = avg_hdr
            elif col == "Final Grade":
                fmt = final_fmt
            ws.write(6, idx, col, fmt)

        for col_idx, col in enumerate(df_final.columns):
            if col.startswith("Average "):
                fmt = avg_data
            elif col == "Final Grade":
                fmt = final_fmt
            else:
                fmt = b_fmt
            for row_offset in range(len(df_final)):
                val = df_final.iloc[row_offset, col_idx]
                ws.write(7 + row_offset, col_idx, "" if pd.isna(val) else val, fmt)

        for idx, col in enumerate(df_final.columns):
            low = col.lower()
            if any(t in low for t in ["first", "last", "name"]):
                ws.set_column(idx, idx, 20)
            elif col.startswith("Average "):
                ws.set_column(idx, idx, 10)
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
    course = st.text_input("Enter class:")
    level = st.text_input("Enter level:")
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            df.columns = df.columns.astype(str)
            output = process_data(df, teacher, subject, course, level)
            st.download_button(
                "Download Organized Gradebook (Excel)", data=output,
                file_name="final_cleaned_gradebook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Processing completed!")
        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
