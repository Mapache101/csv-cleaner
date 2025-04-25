    # Compute category averages using LMS method: (Earned / Possible) * 100
    final_coded = []
    for cat in group_order:
        grp = sorted(groups[cat], key=lambda x: x['seq_num'])
        names = [d['new_name'] for d in grp]
        numeric = df_cleaned[names].apply(lambda x: pd.to_numeric(x, errors='coerce'))

        # Create DataFrame of max points for each assignment
        max_points = pd.DataFrame({d['new_name']: d['max_points'] for d in grp}, index=df_cleaned.index)

        # Only count grades that are not NaN (ignore ungraded assignments)
        has_score = numeric.notna()
        earned = numeric.where(has_score, 0)
        possible = max_points.where(has_score, 0)

        # Sum points earned and possible per student
        sum_earned = earned.sum(axis=1)
        sum_possible = possible.sum(axis=1)

        # Compute raw average as (earned / possible) * 100
        raw_category_average = sum_earned / sum_possible * 100
        raw_category_average = raw_category_average.fillna(0)  # or keep as NaN if preferred

        wt = next((w for k, w in weights.items() if k.lower() == cat.lower()), None)
        avg_col = f"Average {cat}"
        df_cleaned[avg_col] = raw_category_average * wt if wt is not None else raw_category_average
        final_coded.extend(names + [avg_col])
