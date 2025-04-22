import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime
import math  # Import the math module (though ceil isn't used here, round is)

# Weights per category as defined by the Bolivian law
# Using a consistent casing for keys might be slightly safer
weights = {
    "Auto eval": 0.05,
    "TO BE_SER": 0.05,
    "TO DECIDE_DECIDIR": 0.05,
    "TO DO_HACER": 0.40,
    "TO KNOW_SABER": 0.45
}

def process_data(df, teacher, subject, course, level):
    # Updated list of columns to drop from the CSV (if present)
    columns_to_drop = [
        "Nombre de usuario", "Username", "Promedio General", "Term1 - 2024",
        "Term1 - 2024 - AUTO EVAL TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO DECIDE_DECIDIR - Puntuación de categoría",
        "Term1 - 2024 - TO DO_HACER - Puntuación de categoría",
        "Term1 - 2024 - TO KNOW_SABER - Puntuación de categoría",
        "Unique User ID", "Overall", "2025", "Term1 - 2025", "Term2- 2025", "Term3 - 2025"
    ]
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore')

    exclusion_phrases = ["(Count in Grade)", "Category Score", "Ungraded"]
    columns_info = []
    general_columns = []
    columns_to_remove = {"ID de usuario único", "ID de usuario unico"} # Using a set for efficiency

    for i, col in enumerate(df.columns):
        col = str(col) # Ensure header is string
        if col in columns_to_remove or any(phrase in col for phrase in exclusion_phrases):
            continue

        if "Grading Category:" in col:
            m = re.search(r'Grading Category:\s*([^,)]+)', col)
            category = m.group(1).strip() if m else "Unknown"
            base_name = col.split('(')[0].strip()
            new_name = f"{base_name} {category}".strip()
            columns_info.append({
                'original': col, 'new_name': new_name, 'category': category, 'seq_num': i
            })
        else:
            general_columns.append(col)

    name_terms = ["name", "first", "last"]
    name_columns = [col for col in general_columns if any(term in col.lower() for term in name_terms)]
    other_general = [col for col in general_columns if col not in name_columns]
    general_columns_reordered = name_columns + other_general

    sorted_coded = sorted(columns_info, key=lambda x: x['seq_num'])
    new_order = general_columns_reordered + [d['original'] for d in sorted_coded]

    df_cleaned = df[new_order].copy()
    rename_dict = {d['original']: d['new_name'] for d in columns_info}
    df_cleaned.rename(columns=rename_dict, inplace=True)

    groups = {}
    for d in columns_info:
        groups.setdefault(d['category'], []).append(d)
    group_order = sorted(groups.keys(), key=lambda cat: min(d['seq_num'] for d in groups[cat]))

    final_coded_order = []
    # Dictionary to store the UNROUNDED weighted averages (as pandas Series)
    raw_weighted_averages = {}

    for cat in group_order:
        group_sorted = sorted(groups[cat], key=lambda x: x['seq_num'])
        group_names = [d['new_name'] for d in group_sorted]
        avg_col_name = f"Average {cat}" # Final name for the column in Excel

        numeric_group = df_cleaned[group_names].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        raw_avg = numeric_group.mean(axis=1)

        # Find the weight using case-insensitive comparison but fetch the original key casing
        weight_key = next((k for k in weights if k.lower() == cat.lower()), None)
        weight_value = weights.get(weight_key) # Get weight value using the found key

        # Calculate weighted average BUT DO NOT ROUND YET
        if weight_value is not None:
            unrounded_weighted_avg = raw_avg * weight_value
        else:
            # If category not in weights, store the raw average (or handle as needed)
            # For final grade calculation, only weighted categories matter anyway
            unrounded_weighted_avg = raw_avg # Or perhaps 0 or NaN if these shouldn't count

        # Store the unrounded series for final grade calculation
        raw_weighted_averages[avg_col_name] = unrounded_weighted_avg

        # Add the original group columns and the placeholder for the average column to the order
        final_coded_order.extend(group_names)
        final_coded_order.append(avg_col_name) # Add the name for the column that *will* hold the rounded average

    # --- DataFrame Construction and Final Grade Calculation ---
    final_order = general_columns_reordered + final_coded_order
    # Create the final DataFrame structure *before* populating averages/final grade
    df_final = df_cleaned[final_order].copy()

    # Calculate the final grade by summing the UNROUNDED weighted averages
    final_grade_col = "Final Grade"
    def compute_final_grade_sum_first(row_index, raw_weighted_data_dict):
        total = 0.0 # Use float for summation
        valid = False
        # Iterate through the categories defined in weights
        for cat_key, weight in weights.items():
            avg_col_name = f"Average {cat_key}"
            # Check if we have unrounded data for this average column
            if avg_col_name in raw_weighted_data_dict:
                 # Get the unrounded value for the current row
                value = raw_weighted_data_dict[avg_col_name].loc[row_index]
                if pd.notna(value):
                    total += value
                    valid = True
            # If a weighted category was somehow missing from the input, it won't be added
            # which is the desired behavior (can't sum what's not there).
        
        # Round the final sum ONLY HERE
        return int(round(total)) if valid else None # pd.NA might be better than None

    # Apply the new final grade calculation
    # We need to pass the dictionary of unrounded Series
    # Applying row by row is less efficient but direct with the dictionary structure
    final_grades = []
    for index in df_final.index:
         final_grades.append(compute_final_grade_sum_first(index, raw_weighted_averages))
    df_final[final_grade_col] = final_grades


    # Populate the "Average {cat}" columns in the final DataFrame by ROUNDING the stored raw values
    for avg_col_name, unrounded_series in raw_weighted_averages.items():
        if avg_col_name in df_final.columns: # Check if the column exists in the final df
             # Round here for display purposes in the Excel sheet
             df_final[avg_col_name] = unrounded_series.round(0)

    # Replace any occurrence of "Missing" with an empty cell.
    df_final.replace("Missing", "", inplace=True)

    # --- Export to Excel with formatting ---
    output = io.BytesIO()
    with pd.ExcelWriter(
        output,
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    ) as writer:
        # Convert NaN values to empty strings just before writing
        df_final_filled = df_final.fillna('')
        df_final_filled.to_excel(writer, sheet_name='Sheet1', startrow=6, index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Create formats (unchanged)
        header_format = workbook.add_format({'bold': True, 'border': 1, 'rotation': 90, 'shrink': True})
        avg_header_format = workbook.add_format({'bold': True, 'border': 1, 'rotation': 90, 'shrink': True, 'bg_color': '#ADD8E6'}) # Light blue
        avg_data_format = workbook.add_format({'border': 1, 'bg_color': '#ADD8E6'})
        final_grade_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#90EE90'}) # Light green
        border_format = workbook.add_format({'border': 1})

        # Write header info (unchanged)
        worksheet.write('A1', "Teacher:", border_format)
        worksheet.write('B1', teacher, border_format)
        worksheet.write('A2', "Subject:", border_format)
        worksheet.write('B2', subject, border_format)
        worksheet.write('A3', "Class:", border_format)
        worksheet.write('B3', course, border_format)
        worksheet.write('A4', "Level:", border_format)
        worksheet.write('B4', level, border_format)
        # Using current date for timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d") # Changed format slightly for clarity
        worksheet.write('A5', "Generated:", border_format)
        worksheet.write('B5', timestamp, border_format)


        # Write column headers with formatting (unchanged logic)
        for col_num, value in enumerate(df_final.columns):
            if value.startswith("Average "):
                worksheet.write(6, col_num, value, avg_header_format)
            elif value == final_grade_col:
                worksheet.write(6, col_num, value, final_grade_format)
            else:
                worksheet.write(6, col_num, value, header_format)

        # Apply formatting to data cells (using df_final_filled)
        average_columns = [col for col in df_final.columns if col.startswith("Average ")]
        num_rows_data = len(df_final_filled) # Use length of the filled dataframe

        for col_name in df_final_filled.columns:
            col_idx = df_final_filled.columns.get_loc(col_name)
            cell_format = border_format # Default format
            if col_name in average_columns:
                cell_format = avg_data_format
            elif col_name == final_grade_col:
                cell_format = final_grade_format

            # Apply format to the entire column range for data rows
            # Adding 1 to start_row and end_row because xlsxwriter is 0-indexed
            # Start row is 7 (Excel row 8), end row is 7 + num_rows_data - 1
            if num_rows_data > 0:
                worksheet.set_column(col_idx, col_idx, None, cell_format)
            # Note: writing individual cells below is now redundant if set_column format works
            # Let's keep the explicit write for values, but rely on set_column for format
            # The individual write logic seems complex and potentially error-prone, removing it.
            # The df_final_filled.to_excel already wrote the data.

        # Adjust column widths (unchanged)
        for idx, col_name in enumerate(df_final.columns):
            col_width = 5 # Default width
            if any(term in col_name.lower() for term in ["name", "first", "last"]):
                col_width = 25
            elif col_name.startswith("Average"):
                col_width = 7
            elif col_name == final_grade_col:
                col_width = 12
            worksheet.set_column(idx, idx, col_width) # Apply width and existing format

    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Gradebook Organizer")

    # Sidebar instructions (unchanged)
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

    st.title("Griffin CSV to Excel 📊")
    teacher = st.text_input("Enter teacher's name:")
    subject = st.text_input("Enter subject area:")
    course = st.text_input("Enter class:")
    level = st.text_input("Enter level:")
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

    if uploaded_file is not None:
        # Basic check for required fields
        if not all([teacher, subject, course, level]):
             st.warning("Please fill in all the fields (Teacher, Subject, Class, Level).")
        else:
            try:
                # Use 'utf-8' encoding, consider adding 'latin-1' or others if errors occur
                try:
                    df = pd.read_csv(uploaded_file, encoding='utf-8')
                except UnicodeDecodeError:
                    st.warning("UTF-8 decoding failed, trying 'latin-1' encoding.")
                    uploaded_file.seek(0) # Reset file pointer
                    df = pd.read_csv(uploaded_file, encoding='latin-1')

                # Convert all column headers to strings just in case
                df.columns = df.columns.astype(str)

                output_excel = process_data(df, teacher, subject, course, level)

                # Use a dynamic file name including course/subject/date if desired
                file_name = f"Gradebook_{subject}_{course}_{datetime.now().strftime('%Y%m%d')}.xlsx"

                st.download_button(
                    label="Download Organized Gradebook (Excel)",
                    data=output_excel,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Processing completed successfully!")
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
                st.exception(e) # Provides more detailed traceback for debugging

# Standard Python entry point check
if __name__ == "__main__":
    main()