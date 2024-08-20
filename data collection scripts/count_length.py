import pandas as pd

# Load the Excel file
file_path = 'moistmeter_by_category.xlsx'
df = pd.read_excel(file_path)

# Sort the DataFrame by the 'Id' column
df_sorted = df.sort_values(by='ID')

# Save the sorted DataFrame to a new Excel file
sorted_file_path = 'sorted_file.xlsx'
df_sorted.to_excel(sorted_file_path, index=False)

print(f'The rows have been sorted by the "ID" column and saved to {sorted_file_path}')
