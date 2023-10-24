import pandas as pd

df1 = pd.read_excel('sheet1.xlsx')
df2 = pd.read_excel('sheet2.xlsx')

# Merge dataframes based on the columns 'Description' and 'Size'
merged = df1.merge(df2, on=['Description', 'Size'], how='inner')

# Filter rows where 'Description' and 'Size' match, but 'Price' doesn't match
changed_rows = merged[merged['Price_x'] != merged['Price_y']]

# Drop duplicates
unique_changed_rows = changed_rows.drop_duplicates(subset=['Description', 'Size'])

# Alphabetize
unique_changed_rows = unique_changed_rows.sort_values(by=['Description'])

# Select which columns to output
selected_columns = unique_changed_rows[['Description', 'Size', 'Price_x', 'Price_y']]

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter('changes.xlsx', engine='xlsxwriter')

# Write your DataFrame to a sheet named 'changes' and adjust the width of the first column
selected_columns.to_excel(writer, sheet_name='changes', index=False, startrow=0)

# Get the xlsxwriter workbook and worksheet objects for 'changes'
workbook = writer.book
worksheet = writer.sheets['changes']

# Set the 'Description' column width to 40 in 'changes'
worksheet.set_column(0, 0, 40)

# Save the Pandas Excel writer using writer._save()
writer._save()

