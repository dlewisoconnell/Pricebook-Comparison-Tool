import pandas as pd

df1 = pd.read_excel('sheet3.xls')
df2 = pd.read_excel('sheet4.xls')

# Merge dataframes
merged = df1.merge(df2, left_on='PLU/UPC', right_on='PLU/UPC', suffixes=('_sheet1', '_sheet2'), how='inner')

# Filter rows where 'SIZE' matches and 'CASE COST' doesn't match
changed_rows = merged[(merged['SIZE_sheet1'] == merged['SIZE_sheet2']) & (merged['CASE COST_sheet1'] != merged['CASE COST_sheet2'])]

# Drop duplicate rows
changed_rows = changed_rows.drop_duplicates(subset='PLU/UPC')

# Specify the column indices to drop
columns_to_drop_indices = [1, 5, 7, 8, 9, 10, 11] + list(range(12, 22)) + list(range(26, 39))  # Adjusted range

# Drop the specified columns
changed_rows = changed_rows.drop(changed_rows.columns[columns_to_drop_indices], axis=1)

# Drop the "UNIT COST_sheet2" column
changed_rows = changed_rows.drop(columns=['UNIT COST_sheet2', 'SIZE_sheet2', 'WT/CT_sheet2'])

# Check if there are rows that meet the filtering conditions
if not changed_rows.empty:
    # Alphabetize the DataFrame by the "DESCRIPTION_sheet1" column
    unique_changed_rows = changed_rows.sort_values(by='DESCRIPTION_sheet1')

    # Switch columns "PLU/UPC" and "DESCRIPTION_sheet1"
    unique_changed_rows = unique_changed_rows[['DESCRIPTION_sheet1', 'PLU/UPC'] + [col for col in unique_changed_rows.columns if col not in ['PLU/UPC', 'DESCRIPTION_sheet1']]]

    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter('changes.xlsx', engine='xlsxwriter')

    # Write your DataFrame to a sheet named 'changes'
    unique_changed_rows.to_excel(writer, sheet_name='changes', index=False, startrow=0)

    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['changes']

    # Set the width of column A to 40 pixels
    worksheet.set_column(0, 0, 40)

    # Set the width of column B to 15 pixels
    worksheet.set_column(1, 1, 15)

    # Remove the last column
    worksheet.set_column(unique_changed_rows.shape[1], unique_changed_rows.shape[1], None)

    # Save the Pandas Excel writer
    writer._save()
else:
    print("No data matching the filter conditions found.")
