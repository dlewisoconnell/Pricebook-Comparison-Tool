import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Read the data from Excel files
df1 = pd.read_excel('data/week1.xls')
df2 = pd.read_excel('data/week2.xls')

# Merge dataframes
merged = df1.merge(df2, on='PLU/UPC', suffixes=('_week1', '_week2'), how='inner')

# Filter rows where 'SIZE' matches and 'CASE COST' doesn't match
changed_rows = merged[(merged['SIZE_week1'] == merged['SIZE_week2']) & (merged['CASE COST_week1'] != merged['CASE COST_week2'])]

# Drop duplicate rows
changed_rows = changed_rows.drop_duplicates(subset='PLU/UPC')

# Specify the column indices to drop
columns_to_drop_indices = [1, 5, 7, 8, 9, 10, 11] + list(range(12, 22)) + list(range(26, 39))  # Adjusted range

# Drop the specified columns
changed_rows = changed_rows.drop(changed_rows.columns[columns_to_drop_indices], axis=1)

# Drop the "UNIT COST_week2" column
changed_rows = changed_rows.drop(columns=['UNIT COST_week2', 'SIZE_week2', 'WT/CT_week2'])

# Rename the "DESCRIPTION_week1" column to "DESCRIPTION"
changed_rows = changed_rows.rename(columns={'DESCRIPTION_week1': 'DESCRIPTION', 'SIZE_week1': 'SIZE', 'WT/CT_week1': 'WT/CT'})

# Filter rows where 'CASE COST_week2' is greater than 'CASE COST_week1'
filtered_rows = changed_rows[changed_rows['CASE COST_week2'] > changed_rows['CASE COST_week1']].copy()

# Calculate the difference between 'CASE COST_week2' and 'CASE COST_week1'
filtered_rows['Cost_Difference'] = filtered_rows['CASE COST_week2'] - filtered_rows['CASE COST_week1']

# Reorder the columns in the changes file
changed_rows = changed_rows[['DESCRIPTION', 'PLU/UPC'] + [col for col in changed_rows.columns if col not in ['PLU/UPC', 'DESCRIPTION']]]

# Create a pivot table using 'DESCRIPTION' as the index and 'Cost_Difference' as values
pivot_table = filtered_rows.pivot_table(index='DESCRIPTION', values='Cost_Difference', aggfunc='max')

# Sort the pivot table by the maximum difference in descending order
pivot_table = pivot_table.sort_values(by='Cost_Difference', ascending=False)

# Print the sorted pivot table
print("Sorted Pivot Table:")
print(pivot_table)

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter('data/changes.xlsx', engine='xlsxwriter')

# Write your DataFrame to a sheet named 'changes'
changed_rows.to_excel(writer, sheet_name='changes', index=False, startrow=0)

# Write the sorted pivot table to a sheet named 'pivot_table'
pivot_table.to_excel(writer, sheet_name='pivot_table')

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet_changes = writer.sheets['changes']
worksheet_pivot_table = writer.sheets['pivot_table']

# Set the width of columns for 'changes' sheet
for i, col in enumerate(changed_rows.columns):
    column_len = max(changed_rows[col].astype(str).apply(len).max(), len(col))
    worksheet_changes.set_column(i, i, column_len)

# Set the width of columns for 'pivot_table' sheet
for i, col in enumerate(pivot_table.columns):
    column_len = max(pivot_table[col].astype(str).apply(len).max(), len(col))
    worksheet_pivot_table.set_column(i, i, column_len)

# Save the Pandas Excel writer
writer._save()

# Create a bar plot of the top 10 cost differences
top_10_diff = pivot_table.head(10)

plt.figure(figsize=(12, 6))
sns.barplot(x=top_10_diff['Cost_Difference'], y=top_10_diff.index, palette='viridis')
plt.title('Top 10 Cost Increases')
plt.xlabel('Cost Difference')
plt.ylabel('Product Description')
plt.tight_layout()

# Save the plot as an image file
plt.savefig('data/top_10_cost_increases.png')

# Display the plot
plt.show()
