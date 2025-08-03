import win32com.client
import tkinter as tk
from tkinter import filedialog
import os

def select_excel_file():
    root = tk.Tk()
    root.withdraw()  # Skjuler hovedvinduet
    file_path = filedialog.askopenfilename(title="Vælg en Excel-fil", filetypes=[("Excel Files", "*.xls;*.xlsx")])
    return file_path

def copy_excel_file(file_path):
    if not file_path:
        print("Ingen fil valgt.")
        return
    
    excel = win32com.client.Dispatch("Excel.Application")
    
    try:
        workbook = excel.Workbooks.Open(file_path)
        sheet = workbook.Sheets(1)
        sheet.Rows(1).Insert()  # Indsæt en tom række øverst
        
        dir_name, file_name = os.path.split(file_path)
        file_base, file_ext = os.path.splitext(file_name)
        new_file_path = os.path.join(dir_name, f"{file_base}_copy{file_ext}")
        
        workbook.SaveCopyAs(new_file_path)
        print(f"Kopi gemt som: {new_file_path}")
        workbook.Close(False)
    except Exception as e:
        print(f"Fejl under kopiering: {e}")
    finally:
        excel.Quit()
        del excel  # Sørg for at frigøre Excel-objektet

if __name__ == "__main__":
    file_path = select_excel_file()
    copy_excel_file(file_path)