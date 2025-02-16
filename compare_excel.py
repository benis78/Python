import pandas as pd

# Define the file paths
file1_path = r"C:\Coding\Test\Ny mappe\2205-612-001-- - BOM.xlsx"
file2_path = r"C:\Coding\Test\Ny mappe\2205-612-001-A - BOM.xlsx"

# Read the Excel files into pandas DataFrames
try:
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)
except FileNotFoundError:
    print("One or both of the specified files were not found.")
    exit()

# Compare the DataFrames
comparison = df1.compare(df2)

# Output the differences to a new Excel file
comparison.to_excel("differences.xlsx")

print("Comparison complete. Differences saved to differences.xlsx")
