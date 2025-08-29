def process_data(df, teacher, subject, course, level, trimester_choice):
    columns_to_drop = [
        "Nombre de usuario", "Username", "Promedio General",
        "Unique User ID", "2025", "Term3 - 2025"
    ]
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore')

    df.replace("Missing", pd.NA, inplace=True)

    exclusion_phrases = ["(Count in Grade)", "Category Score", "Ungraded"]
    columns_info = []
    general_columns = []
    cols_to_remove = {"ID de usuario único", "ID de usuario unico"}

    for i, col in enumerate(df.columns):
        col = str(col)
        if col in cols_to_remove or any(ph in col for ph in exclusion_phrases):
            continue

        if "Grading Category:" in col:
            m_cat = re.search(r'Grading Category:\s*([^,)]+)', col)
            category = m_cat.group(1).strip() if m_cat else "Unknown"
            m_pts = re.search(r'Max Points:\s*([\d\.]+)', col)
            max_pts = float(m_pts.group(1)) if m_pts else None
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

    name_terms = ["name", "first", "last"]
    name_cols = [c for c in general_columns if any(t in c.lower() for t in name_terms)]
    other_cols = [c for c in general_columns if c not in name_cols]
    
    # --- PROPOSED CHANGE: Rename 'First Name' and 'Last Name' columns ---
    general_reordered = []
    for col in name_cols:
        if col.lower() == 'first name':
            general_reordered.append('Primer Nombre')
        elif col.lower() == 'last name':
            general_reordered.append('Apellidos')
        else:
            general_reordered.append(col)
            
    general_reordered += other_cols
    # --- END PROPOSED CHANGE ---

    sorted_coded = sorted(columns_info, key=lambda x: x['seq_num'])
    new_order = general_reordered + [d['original'] for d in sorted_coded]
    
    # Create a dictionary for renaming the original DataFrame columns
    rename_dict = {
        'First Name': 'Primer Nombre',
        'Last Name': 'Apellidos'
    }
    
    df_cleaned = df.copy()
    df_cleaned.rename(columns=rename_dict, inplace=True, errors='ignore')
    
    df_cleaned.rename({d['original']: d['new_name'] for d in columns_info}, axis=1, inplace=True)

    groups = {}
    for d in columns_info:
        groups.setdefault(d['category'], []).append(d)
    group_order = sorted(groups, key=lambda cat: min(d['seq_num'] for d in groups[cat]))

    final_coded = []
    for cat in group_order:
        grp = sorted(groups[cat], key=lambda x: x['seq_num'])
        names = [d['new_name'] for d in grp]
        
        # --- DYNAMIC LOGIC: Use pre-calculated category score based on trimester choice ---
        category_score_col = f"{trimester_choice} - 2025 - {cat} - Category Score"
        
        raw_avg = pd.Series(dtype='float64')
        if category_score_col in df.columns:
            raw_avg = pd.to_numeric(df[category_score_col], errors='coerce')
        else:
            # Fallback for columns with no space
            category_score_col_no_space = f"{trimester_choice}- 2025 - {cat} - Category Score"
            if category_score_col_no_space in df.columns:
                raw_avg = pd.to_numeric(df[category_score_col_no_space], errors='coerce')
            else:
                numeric = df_cleaned[names].apply(pd.to_numeric, errors='coerce')
                sum_earned = numeric.sum(axis=1, skipna=True)
                max_points_df = pd.DataFrame(index=df_cleaned.index)
                for d in grp:
                    col = d['new_name']
                    max_pts = d['max_points']
                    max_points_df[col] = numeric[col].notna().astype(float) * max_pts
                sum_possible = max_points_df.sum(axis=1, skipna=True)
                raw_avg = (sum_earned / sum_possible) * 100
        
        raw_avg = raw_avg.fillna(0)
        # --- END DYNAMIC LOGIC ---
            
        wt = None
        for key in weights:
            if cat.lower() == key.lower():
                wt = weights[key]
                break
        
        weighted = raw_avg * wt if wt is not None else raw_avg
        avg_col = f"Average {cat}"
        df_cleaned[avg_col] = weighted

        final_coded.extend(names + [avg_col])

    final_order = general_reordered + final_coded
    df_final = df_cleaned[final_order]

    # --- DYNAMIC LOGIC: Use a dynamic column for final grade ---
    final_grade_col = f"{trimester_choice} - 2025"
    final_grade_col_no_space = f"{trimester_choice}- 2025"

    if final_grade_col in df.columns:
        df_final["Final Grade"] = df[final_grade_col]
    elif final_grade_col_no_space in df.columns:
        df_final["Final Grade"] = df[final_grade_col_no_space]
    else:
        df_final["Final Grade"] = pd.NA
    # --- END DYNAMIC LOGIC ---

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter',
                        engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        df_final.to_excel(writer, 'Sheet1', startrow=6, index=False)
        wb = writer.book
        ws = writer.sheets['Sheet1']

        header_fmt = wb.add_format({
            'bold': True,
            'border': 1,
            'rotation': 90,
            'shrink': True,
            'text_wrap': True
        })
        avg_hdr = wb.add_format({
            'bold': True,
            'border': 1,
            'rotation': 90,
            'shrink': True,
            'text_wrap': True,
            'bg_color': '#ADD8E6'
        })
        avg_data = wb.add_format({
            'border': 1,
            'bg_color': '#ADD8E6',
            'num_format': '0'
        })
        final_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#90EE90'})
        b_fmt = wb.add_format({'border': 1})

        ws.write('A1', "Profesor/a:", b_fmt); ws.write('B1', teacher, b_fmt)
        ws.write('A2', "Asignatura:", b_fmt); ws.write('B2', subject, b_fmt)
        ws.write('A3', "Clase:", b_fmt);    ws.write('B3', course, b_fmt)
        ws.write('A4', "Nivel:", b_fmt);    ws.write('B4', level, b_fmt)
        ws.write('A5', datetime.now().strftime("%y-%m-%d"), b_fmt)

        for idx, col in enumerate(df_final.columns):
            fmt = header_fmt
            if col.startswith("Average "):
                fmt = avg_hdr
            elif col == "Final Grade":
                fmt = final_fmt
            ws.write(6, idx, col, fmt)

        avg_cols = {c for c in df_final.columns if c.startswith("Average ")}
        for col_idx, col in enumerate(df_final.columns):
            fmt = avg_data if col in avg_cols else final_fmt if col == "Final Grade" else b_fmt
            for row_offset in range(len(df_final)):
                val = df_final.iloc[row_offset, col_idx]
                excel_row = 7 + row_offset
                ws.write(excel_row, col_idx, "" if pd.isna(val) else val, fmt)

        name_terms = ["name", "first", "last"]
        for idx, col in enumerate(df_final.columns):
            if any(t in col.lower() for t in name_terms):
                ws.set_column(idx, idx, 25)
            elif col.startswith("Average "):
                ws.set_column(idx, idx, 7)
            elif col == "Final Grade":
                ws.set_column(idx, idx, 12)
            else:
                ws.set_column(idx, idx, 10)

    return output
