from Compare_1 import compare_excel
import tkinter as tk
from tkinter import filedialog, messagebox
import os.path
import pandas as pd
import re

SheetName='PARTLIST'
root = tk.Tk()
root.withdraw()
t='Open Excel PARTLIST file'


# path1=filedialog.askopenfilename(title=t.upper(), filetypes=[('Excel files','.xlsx .xls')], initialdir='C:\\Working Folder\\Designs\\5-Projects')
# print(path1)
    
#path1="C:\\Working Folder\\Designs\\5-Projects\\2500 - Aabenraa\\Area 02 - Mixing Tank Area\\2500-402-001-A - PARTLIST.xlsx"

path1=filedialog.askopenfilename(title=t.upper(), filetypes=[('Excel files','.xlsx .xls')], initialdir='C:\\Working Folder\\Designs\\5-Projects')
path1base=(os.path.basename(path1))
path1dir=(os.path.dirname(path1))
if path1 == '':
    os._exit(1)


# get excel sheet name
xls1 = pd.ExcelFile(path1)
sheet1 = xls1.sheet_names

if sheet1[0] != SheetName:
    print(sheet1[0])
    tk.messagebox.showinfo('Info','Must be a PARTLIST')
    #os._exit(1)
    
#get revision letter
rlo = re.compile(r'-\D - ')
ro=rlo.findall(path1base)[0][1:2]
listo=[]



#Check for uniqe Filename
dfo=pd.read_excel(path1)
dfo=pd.DataFrame(dfo,columns=['Filename'],)
# dfo[dfo.duplicated(subset="Filename",keep=False)]
dup=dfo[dfo.duplicated(keep=False)]
print(dup)     






t='Open Excel PARTLIST file'
path2=filedialog.askopenfilename(title=t.upper(), filetypes=[('Excel files','.xlsx .xls')], initialdir=path1dir)
path2base=(os.path.basename(path2))
path2dir=(os.path.dirname(path2))
if path2 == '':
    os._exit(1)
# get excel sheet name
xls2 = pd.ExcelFile(path2)
sheet2 = xls2.sheet_names

rln = re.compile(r'-\D - ')
rn=rln.findall(path2base)[0][1:2]

# check if sheets name are the same
if sheet1 == sheet2:
    sheetname = sheet1[0]
    if sheetname == 'PARTLIST':
        key_column = 'Filename'
    # if sheetname == 'BOM':
    #     key_column = 'Item'
elif sheet1[0] != sheet2[0]:
    tk.messagebox.showinfo('Error Sheet name','You must compare Partlist with Partlist! Use Export ilogic in Inventor, Try again')
    os._exit(1)

output_path = path2dir+'//Compare '+path2base[0:11]+' ('+ro+')-('+rn+').xlsx'

compare_excel(path1, path2, output_path, sheetname, key_column)#, skiprows=1)


