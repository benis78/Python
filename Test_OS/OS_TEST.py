import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

t='Open folder'
folder_path=filedialog.askdirectory() #(title=t.upper(), filetypes=[('Excel files','.xlsx .xls')], initialdir='C:\\Working Folder\\Designs\\5-Projects')
#print(file_path)
#file_path='D:/Dropbox/Coding/BOM.xlsx'
fPath=(os.path.dirname(folder_path))
if folder_path == '':
    os._exit(1)

print (fPath)
print (os.path.relpath(fPath))
# print('The current directory is: %s' % os.path.curdir)
# # dest_path=''
# # destPath=(os.path.abspath(dest_path))

targetExtensions = ['.pdf']
excludeExtensions= '_FOR REVIEW.pdf'

for root, dirs, files in os.walk(fPath):           # Vi looper igennem vores directory
    if 'Old' in dirs:
        dirs.remove('Old')
    if 'old' in dirs:
        dirs.remove('old')
    if 'OLD' in dirs:
        dirs.remove('OLD')
    for name in files:                              # Vi looper igennem hver fil i den aktuelle mappe
        print(name)
        #Hver fil bliver så sammenlignet med hver værdi i targetFileNames med hver af de mulige extensions:
        # for targetName in targetFileNames:
        #     for targetExtension in targetExtensions:
        #         if name.startswith(targetName) and name.endswith(targetExtension): 
        #             if not name.endswith(excludeExtensions):   
                    # Når vi finder et match, bliver stien til den aktuelle fil tilføjet på den passende position i variablen fileDestinations:
                        #fileDestinations[targetFileNames.index(targetName)]
                        #fileDestinations[targetFileNames.index(targetName)].append(os.path.join(root,name))
                        # idx=-1
                        # while True:
                        #     try:
                        #         idx = targetFileNames.index(targetName, idx+1)
                        #         fileDestinations[idx].append(os.path.join(root,name))
                        #     except ValueError:
                        #         break
                       
                        #fileDestinations[ti].append(os.path.join(root,name))
                        #print(fileDestinations, end= '\n')