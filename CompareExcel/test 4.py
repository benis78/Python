import pandas as pd
from tkinter import filedialog

def compare_excel_files(file1, file2, output_file):
    # Indlæs Excel-filerne
    df1 = pd.read_excel(file1, sheet_name="PARTLIST")
    df2 = pd.read_excel(file2, sheet_name="PARTLIST")
    
    # Definer nøglekolonner
    key_columns = ["Part Number", "Description"]
    
    # Merge for at finde ændrede rækker
    merged = df1.merge(df2, on=key_columns, how="outer", indicator=True, suffixes=("_old", "_new"))
    
    # Find ændrede rækker (kun hvis der er forskelle i andre kolonner)
    changed_columns = [col for col in df1.columns if col not in key_columns and col in df2.columns]
    changed = merged[(merged["_merge"] == "both") & (merged.apply(lambda row: any(row.get(col + "_old", None) != row.get(col + "_new", None) for col in changed_columns if (col + "_old" in merged.columns and col + "_new" in merged.columns)), axis=1))]
    changed = changed.drop(columns=["_merge"])
    
    # Find nye rækker
    new_rows = merged[merged["_merge"] == "right_only"].drop(columns=["_merge"])
    
    # Find fjernede rækker
    removed_rows = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])
    
    # Gem til en ny Excel-fil med tre faner
    with pd.ExcelWriter(output_file) as writer:
        changed.to_excel(writer, sheet_name="Changed", index=False)
        new_rows.to_excel(writer, sheet_name="New", index=False)
        removed_rows.to_excel(writer, sheet_name="Removed", index=False)
    
    print(f"Sammenligning færdig! Resultater gemt i {output_file}")

# Vælg filer via fil-dialog
file1 = filedialog.askopenfilename(title="Vælg den første Excel-fil", filetypes=[("Excel files", "*.xlsx *.xls")])
file2 = filedialog.askopenfilename(title="Vælg den anden Excel-fil", filetypes=[("Excel files", "*.xlsx *.xls")])
output_file = "Compare.xlsx"

if file1 and file2:
    compare_excel_files(file1, file2, output_file)
else:
    print("Filer ikke valgt. Afbryder sammenligning.")
