import pandas as pd
import PyPDF2
import os
import glob

# Indlæs Excel-filen
excel_fil = "BasicComponentMerge/BasicComponent.xlsx"  # Sørg for, at filnavnet er korrekt
sheet_name = "Ark1"  # Opdater hvis nødvendigt
kolonne_index = 1  # Anden kolonne (0-baseret indeks)

# Angiv undermappen, hvor PDF-filerne findes
pdf_mappe = "BasicComponentMerge"

# Læs Excel-filen og fjern tomme værdier
df = pd.read_excel(excel_fil, sheet_name=sheet_name)
pdf_søgning = df.iloc[1:, kolonne_index].dropna().astype(str).str.strip().tolist()  # Fjern mellemrum

# Find alle PDF-filer i undermappen
alle_pdf_filer = [os.path.join(pdf_mappe, f) for f in os.listdir(pdf_mappe) if f.endswith(".pdf")]

# Udskriv de PDF-filer, der findes i undermappen
print("PDF-filer i undermappen:")
for pdf in alle_pdf_filer:
    print(f"- {pdf}")

# Udskriv de søgetekster, der bruges
print("\nSøgetekster fra Excel:")
for søgetekst in pdf_søgning:
    print(f"- {søgetekst}")

# Initialiser en PDF merger
merger = PyPDF2.PdfMerger()

# Gennemgå søgeteksterne og find PDF-filer, der indeholder søgeteksten et sted i navnet
for søgetekst in pdf_søgning:
    match_fundet = False
    for pdf in alle_pdf_filer:
        if søgetekst in os.path.basename(pdf):  # Matcher søgetekst i filnavnet (uden sti)
            merger.append(pdf)
            print(f"Tilføjer: {pdf} (match med '{søgetekst}')")
            match_fundet = True
            break  # Stop efter første match
    if not match_fundet:
        print(f"Advarsel: Ingen PDF fundet for '{søgetekst}', udelades.")

# Gem den flettede PDF
output_fil = "samlet_fil.pdf"
merger.write(output_fil)
merger.close()

print(f"Fletning færdig! Den nye PDF er gemt som: {output_fil}")
