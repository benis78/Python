import glob, os
os.chdir(r'D:\Dropbox\VizCon\03-Teknisk dokumentation\Vitek (1)\Projekt\Flexicon\Flexicon Svinger')
for file in glob.glob("*.pdf"):
    print (file)