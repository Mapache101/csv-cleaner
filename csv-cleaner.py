import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime
import math

# ——— 1) Make the weights dict use the exact category strings ———
weights = {
    "AUTO EVAL":        0.05,
    "TO BE_SER":        0.05,
    "TO DECIDE_DECIDIR":0.05,
    "TO DO_HACER":      0.40,
    "TO KNOW_SABER":    0.45
}

def custom_round(value):
    """Round half-up: .5 and above rounds up."""
    return math.floor(value + 0.5)

def process_data(df, teacher, subject, course, level):
    # ——— 2) Drop Schoology’s pre-computed “Category Score” columns ———
    to_drop = [
        "Nombre de usuario", "Username", "Promedio General",
        "Term1 - 2024", "Term1 - 2024 - AUTO EVAL - Category Score",
        "Term1 - 2024 - TO BE_SER - Category Score",
        "Term1 - 2024 - TO DECIDE_DECIDIR - Category Score",
        "Term1 - 2024 - TO DO_HACER - Category Score",
        "Term1 - 2024 - TO KNOW_SABER - Category Score",
        "Unique User ID", "Overall", "2025", "Term1 - 2025",
        "Term2- 2025", "Term3 - 2025"
    ]
    df = df.drop(columns=to_drop, errors='ignore')

    # Treat “Missing” as blank → NaN
    df.replace("Missing", pd.NA, inplace=True)

    exclusion_phrases = ["(Count in Grade)", "Category Score", "Ungraded"]
    columns_info = []
    general_columns = []
    cols_to_remove = {"ID de usuario único", "ID de usuario unico"}

    # ——— 3) Identify every “Grading Category: X” column ———
    for i, col in enumerate(df.columns):
        col = str(col)
        if col in cols_to_remove or any(ph in col for ph in exclusion_phrases):
            continue

        if "Grading Category:" in col:
            # Extract the category name
            m_cat = re.search(r'Grading Category:\s*([^,)]+)', col)
            category = m_cat.group(1).strip() if m_cat else "Unknown"

            # Extract the max-points
            m_pts = re.search(r'Max Points:\s*([\d\.]+)', col)
            max_pts = float(m_pts.group(1)) if m_pts else 0.0

            # Build a clean column name
            base = col.split('(')[0].strip()
            new_name = f"{base} {category}"

            columns_info.append({
                'original':   col,
                'new_name':   new_name,
                'category':   category,
                'seq_num':    i,
                'max_points': max_pts
            })
        else:
            general_columns.append(col)

    # ——— 4) Pull out “First Name” / “Last Name” first ———
    name_terms = ["name", "first", "last"]
    name_cols  = [c for c in general_columns if any(t in c.lower() for t in name_terms)]

    # Sort assignments in their original sequence
    sorted_asmts = sorted(columns_info, key=lambda d: d['seq_num'])
    df_clean = df[name_cols + [d['original'] for d in sorted_asmts]].copy()
    df_clean.rename({d['original']: d['new_name'] for d in columns_info},
                    axis=1, inplace=True)

    # ——— 5) Group by category & compute each category’s weighted % ———
    groups = {}
    for d in columns_info:
        groups.setdefault(d['category'], []).append(d)

    # For each category in the WEIGHTS order:
    avg_cols = []
    for cat in weights:
        if cat not in groups:
            continue
        items = sorted(groups[cat], key=lambda d: d['seq_num'])
        names = [d['new_name'] for d in items]

        # Numeric scores, blanks → 0
        raw = (df_clean[names]
               .apply(pd.to_numeric, errors='coerce')
               .fillna(0))

        total_earned   = raw.sum(axis=1)
        total_possible = sum(d['max_points'] for d in items) or 1.0
        pct            = total_earned / total_possible * 100
        contrib        = pct * weights[cat]

        avg_name = f"Average {cat}"
        df_clean[avg_name] = contrib
        avg_cols.append((cat, names, avg_name))

    # ——— 6) Final Grade = sum of weighted contributions, rounded half-up ———
    df_clean["Final Grade"] = (
        df_clean[[avg for (_,_,avg) in avg_cols]]
        .sum(axis=1)
        .apply(custom_round)
    )

    # ——— 7) Reorder to: Name, Last Name, [per-category scores + Average], …, Final Grade ———
    ordered = []
    ordered.extend(name_cols)
    for cat, names, avg in avg_cols:
        ordered.extend(names)   # all raw scores in that category
        ordered.append(avg)     # then its Average
    ordered.append("Final Grade")

    df_final = df_clean[ordered]

    # ——— 8) Export to Excel exactly as before ———
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter',
                        engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:

        df_final.to_excel(writer, 'Sheet1', startrow=6, index=False)
        wb = writer.book
        ws = writer.sheets['Sheet1']

        # Write your header info
        b_fmt = wb.add_format({'border': 1})
        ws.write('A1', "Teacher:", b_fmt); ws.write('B1', teacher, b_fmt)
        ws.write('A2', "Subject:", b_fmt); ws.write('B2', subject, b_fmt)
        ws.write('A3', "Class:",   b_fmt); ws.write('B3', course,  b_fmt)
        ws.write('A4', "Level:",   b_fmt); ws.write('B4', level,   b_fmt)
        ws.write('A5', datetime.now().strftime("%y-%m-%d"), b_fmt)

        # Formats for headers & data
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
        for col_idx, col in enumerate(df_final.columns):
            if col == "Final Grade":
                fmt = final_fmt
            elif col.startswith("Average "):
                fmt = avg_hdr
            else:
                fmt = header_fmt
            ws.write(6, col_idx, col, fmt)

        # Write data rows
        for col_idx, col in enumerate(df_final.columns):
            if col == "Final Grade":
                fmt = final_fmt
            elif col.startswith("Average "):
                fmt = avg_data
            else:
                fmt = b_fmt
            for row in range(len(df_final)):
                val = df_final.iloc[row, col_idx]
                ws.write(7 + row, col_idx,
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

# ——— The rest of your Streamlit main() stays unchanged ———
def main():
    st.set_page_config(page_title="Gradebook Organizer")
    # … your sidebar + inputs …
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])
    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
        df.columns = df.columns.astype(str)
        out = process_data(df, teacher, subject, course, level)
        st.download_button("Download Organized Gradebook (Excel)",
                            data=out,
                            file_name="gradebook.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("Done!")

if __name__ == "__main__":
    main()
