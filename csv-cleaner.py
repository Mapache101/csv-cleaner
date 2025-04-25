def process_data(df, groups):
    df_cleaned = df.copy()
    output_columns = []
    all_scores = []

    for group in groups:
        category_name = group['category']
        assignments = group['assignments']
        max_points = group['max_points']
        weight = group['weight']
        avg_col = group['new_name']

        # Collect existing assignment columns
        existing_assignments = [a for a in assignments if a in df_cleaned.columns]

        if not existing_assignments:
            continue

        # Convert to numeric
        df_cleaned[existing_assignments] = df_cleaned[existing_assignments].apply(pd.to_numeric, errors='coerce')

        # Calculate category total points (earned / max)
        earned = df_cleaned[existing_assignments].sum(axis=1)
        total_possible = sum([max_points[a] for a in existing_assignments if a in max_points and pd.notnull(df_cleaned[a]).any()])

        if total_possible == 0:
            df_cleaned[avg_col] = 0
        else:
            df_cleaned[avg_col] = (earned / total_possible) * 100

        # Maintain column group structure: assignments + category average
        output_columns.extend(existing_assignments + [avg_col])
        all_scores.append((avg_col, weight))

    # Weighted average of all category scores
    total_weight = sum(w for _, w in all_scores)
    if total_weight > 0:
        df_cleaned['Weighted Average'] = sum(df_cleaned[col] * (w / total_weight) for col, w in all_scores)
        output_columns.append('Weighted Average')

    return df_cleaned[output_columns]
