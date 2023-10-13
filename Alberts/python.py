import pandas as pd
df1 = pd.read_excel('sheet1.xlsx')
df2 = pd.read_excel('sheet2.xlsx')

# Merge dataframes
merged = df1.merge(df2, on=['Product Description', 'Size'], how='inner')

# Filter rows where 'Price' doesn't match
changed_rows = merged[merged['Price_x'] != merged['Price_y']]

# Drop rows with missing values as 'Product Description'
changed_rows = changed_rows.dropna(subset=['Product Description'])

# Show only descriptions beginning with 'ORG'
org_rows = changed_rows[changed_rows['Product Description'].str.startswith('ORG')]

# Drop duplicates
org_rows = org_rows.drop_duplicates(subset=['Product Description', 'Size'])

# Sort alphabetically
org_rows = org_rows.sort_values(by='Product Description')

# Select columns to output to new file
selected_columns = org_rows[['Product Description', 'Size', 'Price_x', 'Price_y']]

# Output to excel file
selected_columns.to_excel('changes.xlsx', index=False)

