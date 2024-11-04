import streamlit as st
import pandas as pd
import numpy as np

# Display the app header
st.title("TSI Benchmarks Validator")

# File upload widgets
uploaded_excel = st.file_uploader("Upload the TSI Benchmarks Excel File (.xlsm)", type="xlsm")
uploaded_csv = st.file_uploader("Upload the Calculated Benchmarks CSV File", type="csv")

if uploaded_excel and uploaded_csv:
    # Read the Excel file and parse the specific sheet
    excel = pd.ExcelFile(uploaded_excel)
    df = excel.parse('Unseen Extract')
    df = df.loc[df.survey_id.notnull()]

    # Columns to subset and clean
    # benchmarks = df.columns[16:50] might not be consistent

    benchmarks = ['General Manager', 'Property Manager', 'CRM / TSM',
       'Facilities Manager', 'Senior Facilities Manager',
       'Management Staff (All)', 'Standard Management Team Staff',
       'Service Requests', 'Fitout Management', 'Rental Billing Management',
       'Management Team Average', 'Access Control Management',
       'Air Conditioning', 'Food and Beverage Offering',
       'Car Park Mgnt (Internal)', 'Car Park Mgnt (External)',
       'Car Park Operator (Name)', 'Cleaning - Offices', 'Cleaning - Bathroom',
       'Cleaning - Common', 'Cleaning - Average', 'Concierge',
       'Emergency Management & Safety', 'End of Trip', 'Lifts', 'Presentation',
       'Security', 'ESG Communication', 'Warden Training',
       'Building Services Average', 'Property Performance', 'TSI Metro',
       'Management Staff (All) Index', 'Standard Management Team Staff Index']
    
    blacklist = set(['Car Park Operator (Name)', 'TSI Metro', 'Management Staff (All)', 'Standard Management Team Staff'])
    COLS = list(set(benchmarks) - blacklist)

    ORDER = [
        'Property Performance', 'Management Team Average', 'Building Services Average', 'Management Staff (All) Index', 'Standard Management Team Staff Index',
        'General Manager', 'Property Manager', 'CRM / TSM', 'Facilities Manager', 'Senior Facilities Manager', 
        'Service Requests', 'Fitout Management', 'Rental Billing Management', 'Access Control Management', 'Air Conditioning', 'Food and Beverage Offering', 
        'Car Park Mgnt (Internal)', 'Car Park Mgnt (External)', 
        'Cleaning - Offices', 'Cleaning - Bathroom', 'Cleaning - Common', 'Cleaning - Average', 'Concierge', 'Emergency Management & Safety',
        'End of Trip', 'Lifts', 'Presentation', 'Security', 'ESG Communication', 'Warden Training'
    ]

    # Function to reorder columns
    def reorder_list(original_list, desired_order):
        original_set = set(original_list)
        desired_set = set(desired_order)
        missing_from_original = desired_set - original_set
        missing_from_desired = original_set - desired_set

        if missing_from_original or missing_from_desired:
            error_msg = []
            if missing_from_original:
                error_msg.append(f"Elements missing from original list: {sorted(missing_from_original)}")
            if missing_from_desired:
                error_msg.append(f"Elements missing from desired order: {sorted(missing_from_desired)}")
            raise ValueError("\n".join(error_msg))

        order_map = {item: index for index, item in enumerate(desired_order)}
        return sorted(original_list, key=lambda x: order_map[x])

    # Reorder columns
    ORDERED_COLS = reorder_list(COLS, ORDER)
    for col in ORDERED_COLS:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Read the CSV file
    df_snowflake = pd.read_csv(uploaded_csv)
    qualtrics = df_snowflake.loc[df_snowflake.survey_id.notnull()]

    # Merge the two DataFrames
    merge = qualtrics.merge(df, left_on='survey_id', right_on='survey_id', how='outer', suffixes=('_qualtrics', '_tsi'))

    # Validation function
    def validate_survey_results(df1, df2, survey_id, candidate_columns, suffixes=('_qualtrics', '_tsi')):
        merged_df = df1.merge(df2, left_on=survey_id, right_on=survey_id, suffixes=suffixes)
        result_rows = []

        for _, row in merged_df.iterrows():
            for suffix in suffixes:
                result_row = {survey_id: row[survey_id], 'source': suffix}
                for col in candidate_columns:
                    result_row[col] = row[f"{col}{suffix}"]
                result_rows.append(result_row)

            diff_row = {survey_id: row[survey_id], 'source': '_difference'}
            for col in candidate_columns:
                diff_row[col] = abs(row[f"{col}{suffixes[0]}"] - row[f"{col}{suffixes[1]}"])
            result_rows.append(diff_row)

        return pd.DataFrame(result_rows)

    # Apply validation
    r = validate_survey_results(qualtrics, df, 'survey_id', ORDERED_COLS)

    # Filter significant differences
    def filter_large_differences(df, candidate_columns, thresh):
        diff_rows = df[df['source'] == '_difference']
        mask = diff_rows[candidate_columns].gt(thresh).any(axis=1)
        return diff_rows[mask]

    significant_differences = filter_large_differences(r, ORDERED_COLS, 0.1)
    ID = ['survey_id', 'source']
    final = r.loc[r.survey_id.isin(significant_differences.survey_id)][ID + ORDERED_COLS]

    # Display the DataFrame and allow download
    st.subheader("Significant Differences")
    st.dataframe(final)

    # Option to download CSV
    csv = final.to_csv(index=False).encode('utf-8')
    st.download_button("Download Significant Differences as CSV", csv, "significant_differences.csv")

else:
    st.warning("Please upload both the Excel and CSV files to proceed.")
