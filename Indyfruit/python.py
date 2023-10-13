import pandas as pd
df1 = pd.read_excel('sheet1.xls')
df2 = pd.read_excel('sheet2.xls')

# Merge dataframes
merged = df1.merge(df2, on='DESCRIPTION', suffixes=('_sheet1', '_sheet2'), how='inner')

# Filter rows where 'CASE COST' doesn't match
changed_rows = merged[merged['CASE COST_sheet1'] != merged['CASE COST_sheet2']]

# Drop duplicates
unique_changed_rows = changed_rows.drop_duplicates(subset=['DESCRIPTION', 'SIZE_sheet1', 'SIZE_sheet2'])

# Convert 'SIZE_sheet2' to numeric
unique_changed_rows.loc[:, 'SIZE_sheet2'] = pd.to_numeric(unique_changed_rows['SIZE_sheet2'], errors='coerce')

# Alphabetize
unique_changed_rows = unique_changed_rows.sort_values(by='DESCRIPTION')

# Select which columns to output
selected_columns = unique_changed_rows[['DESCRIPTION', 'SIZE_sheet1', 'WT/CT_sheet1', 'CASE COST_sheet1', 'SIZE_sheet2', 'WT/CT_sheet2', 'CASE COST_sheet2']]

# Output data to a new excel sheet
selected_columns.to_excel('changes.xlsx', index=False)
