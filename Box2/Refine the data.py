import pandas as pd

# Load the data from an Excel file
file_path = 'Reports/Darvas_Box_Trade_Combined_Report.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Step 1: Sort the data by 'Stock Symbol', 'Exit Date', and 'Entry Date' (oldest first)
df_sorted = df.sort_values(by=['Stock Symbol', 'Exit Date', 'Entry Date'])

# Step 2: Drop duplicates by keeping the oldest 'Entry Date' for each 'Stock Symbol' and 'Exit Date' combination
df_filtered = df_sorted.drop_duplicates(subset=['Stock Symbol', 'Exit Date'], keep='first')

# Step 3: Sort the filtered data by 'Stock Symbol' and 'Entry Date' in descending order
df_filtered_sorted = df_filtered.sort_values(by=['Stock Symbol', 'Entry Date'], ascending=[True, False])

# Save the refined data to a new Excel file (optional) 
output_path = 'Reports/refined_output.xlsx'
df_filtered_sorted.to_excel(output_path, index=False)

print("Data refinement and sorting are complete. Refined data saved to:", output_path)
