import os, openpyxl, time
import tkinter as tk
from tkinter import filedialog

t='Open Excel file'
file_path=filedialog.askopenfilename(title=t.upper(), filetypes=[('Excel files','.xlsx .xls')])
efile=(os.path.basename(file_path))
ePath=(os.path.dirname(file_path))
eBOM=(os.path.abspath(file_path))
if file_path == '':
    os._exit(1)

os.chdir(ePath)
wb=openpyxl.load_workbook(eBOM)
sheet = wb['Sheet1']

dPath=(r'\\192.168.170.9\Bigadan\Bigadan\1_PROJEKTER\2205 Solrød Biogas\220511 Solrød ombygning 2019\6.0 Drawings')

fileStart=('2','0000-3')
i=1
for root, dirs, files in os.walk(dPath):           # Vi looper igennem vores directory
    if 'Old' in dirs:
        dirs.remove('Old')
    for name in files:
        if name.endswith('.pdf') and name.startswith(fileStart):
            sheet.cell(row=i, column=1).value = name
            i+=1
            
wb.save(efile)