import os
import pandas as pd
cwd = os.path.abspath('C:\\Users\\Jesper\\Desktop\\Ny mappe (2)') 
files = os.listdir(cwd)  

## Method 1 gets the first sheet of a given file
df = pd.DataFrame()
for file in files:
    if file.endswith('.xlsx'):
        df = df.append(pd.read_excel(file), ignore_index=True) 
df.head() 
df.to_excel('total_sales.xlsx')



## Method 2 gets all sheets of a given file
df_total = pd.DataFrame()
for file in files:                         # loop through Excel files
    if file.endswith('.xlsx'):
        excel_file = pd.ExcelFile(file)
        sheets = excel_file.sheet_names
        for sheet in sheets:               # loop through sheets inside an Excel file
            df = excel_file.parse(sheet_name = sheet)
            df_total = df_total.append(df)
df_total.to_excel('combined_file.xlsx')


#P_o=Path(r'C:\Users\Jesper\Desktop\Ny mappe (2)\P&ID Equipment 20210806-08.38.xlsx')
#P_n=Path(r'C:\Users\Jesper\Desktop\Ny mappe (2)\P&ID Equipment 20211026-13.32.xls')

#excel_diff(P_o, P_n)
  
#C:\\Users\\Jesper\\Desktop\\Ny mappe (2)\\P&ID Equipment 20210806-08.38.xlsx
#C:\\Users\\Jesper\\Desktop\\Ny mappe (2)\\P&ID Equipment 20211026-13.32.xls
  
