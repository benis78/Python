"""
NØGLEREGLER OG PROCESSER FOR EXCEL BOM HÅNDTERING

1. DATAHÅNDTERING OG KATEGORISERING
--------------------------------
- Brug pandas til al datahåndtering og analyse
- Anvend kategoriske datatyper for effektiv hukommelsesudnyttelse
- Implementér feature engineering for kategorisering
- Brug groupby og pivot_table til aggregering

2. PROCESRÆKKEFØLGE
----------------
TRIN 1: Data Indlæsning og Validering
    - Indlæs Excel BOM med pandas
    - Valider datatyper og struktur
    - Tjek for manglende værdier
    - Konverter relevante kolonner til kategoriske datatyper

TRIN 2: Kategorisering
    - Anvend vektoriserede operationer for kategorisering
    - Gruppér data baseret på hierarkisk struktur
    - Implementér piping kategorisering først
    - Derefter normal kategorisering

TRIN 3: Filhåndtering
    - Scan kildemappen én gang
    - Cache resultaterne for bedre performance
    - Brug concurrent.futures for parallel filscanning
    - Valider fileksistens før kopiering

TRIN 4: Excel Generering
    - Brug openpyxl til Excel manipulation
    - Implementér formatering én gang per ark
    - Gruppér data effektivt
    - Sikr korrekt rækkefølge i hierarkisk struktur

3. FEJLHÅNDTERING
--------------
- Implementér try-except blokke for alle eksterne operationer
- Valider data ved hver transformationstrin
- Log fejl og advarsler til debug.txt
- Giv brugeren klare fejlbeskeder

4. PERFORMANCE OPTIMERING
---------------------
- Brug vektoriserede operationer hvor muligt
- Implementér caching for gentagne operationer
- Minimer disk I/O operationer
- Brug parallel processing for tunge operationer

5. DATAVALIDERING
-------------
Pre-processing validering:
    - Excel struktur
    - Påkrævede kolonner
    - Datatyper
    - Manglende værdier

Post-processing validering:
    - Kategorisering komplethed
    - Filkopiering success
    - Excel ark integritet
    - Gruppering korrekthed

6. KODESTRUKTUR
------------
- Modulær opbygning med veldefinerede funktioner
- Klar separation af ansvar
- Dokumenterede funktioner med docstrings
- Følg PEP 8 standarder

7. LOGGING OG DEBUGGING
-------------------
- Detaljeret logging til debug.txt
- Klare statusmeddelelser til bruger
- Performance metrics tracking
- Fejlsporing med stack traces

8. FILHÅNDTERING
------------
Kildefiler:
    - PDF og DWG filer i netværksmappe
    - Excel BOM fil
    - Kategoriseringsfiler (categories.txt, piping_categories.txt)

Destinationsfiler:
    - Kategoriserede mapper med kopierede filer
    - Ny Excel fil med kategoriserede faneblade
    - Debug log fil

9. BRUGERINTERAKTION
----------------
- Klar fremgangsindikator
- Informative fejlmeddelelser
- Mulighed for at afbryde langvarige processer
- Bekræftelsesdialog ved kritiske operationer
"""

# Rækkefølge for processing
PROCESSING_STEPS = [
    "Opretter kopi af Excel fil",
    "Identificerer kolonner",
    "Håndterer Inseparable rækker",
    "Håndterer revisioner",
    "Beregner Total QTY",
    "Indskyder linje i række 2 for at få part number med i scanningen",
    "Fjerner Inseparable children linjer",
    "Scanner for alle rør efter Piping_categories.txt",
    "Scanner alle andre part numbers for kategori (undtagen piping items)",
    "Formaterer Excel ark",
    "Opretter kategori faner",
    "Opretter mapper",
    "Kopierer filer efter kategorisering"
]

# Kolonner der skal identificeres
REQUIRED_COLUMNS = [
    "ITEM",
    "PART NUMBER", 
    "REV",
    "BOM STRUCTURE",
    "QTY",
    "DRAWINGS",
    "TOTAL QTY",
    "DESCRIPTION"
]

# Standardværdier for indskudt række
INSERTED_ROW = {
    "ITEM": "0",
    "PART NUMBER": None,  # Udfyldes dynamisk baseret på filnavn
    "REV": "-",
    "BOM STRUCTURE": "INSEPARABLE",
    "QTY": "1",
    "TOTAL QTY": "1",
    "DESCRIPTION": "Area Layout Drawing"
}


# Netværkssti konfiguration
NETWORK_PATH = r'C:\Working Folder\Designs\5-Projects\4003 - Nurmo Bioenergia\Area 05 - Storage Area\Equipment\BOM\Test\Files'

# Piping kategori instruktioner
PIPING_RULES = """
Kategorisering skal følge denne rækkefølge:
1. Gruppér først alle rækker efter deres item numbers (f.eks. 2 (Parent), 2.1(Child), 2.1.1(Child), 2.2(Child), 3(Parent), 3.1(Child) osv. hører sammen)
2. For hver gruppe:
   - Tjek parent item's part number (f.eks. item "2")
   - Hvis parent item's part number matcher en piping kategori (f.eks. BM):
     * Flyt HELE gruppen (parent + alle under-items) til den pågældende piping kategori
     * Markér alle disse items som "allerede kategoriseret"
3. For alle resterende items der IKKE er markeret som "allerede kategoriseret":
   - Kategoriser dem efter categories.txt
   - Undlad at kategorisere items der allerede er i en piping kategori

Eksempel:
Item    Part Number
2       4003-02.1-BM003-01    <- Parent med BM
2.1     0000-701-051          <- Under-item (skal med i BM kategori)
2.1.1   0000-703-040          <- Under-item (skal med i BM kategori)
2.2     0000-703-020          <- Under-item (skal med i BM kategori)

3       4003-02.1-E01         <- Nyt parent item (kategoriseres normalt)
3.1     0000-301-010          <- Under-item (kategoriseres normalt)



# Kolonne placeringer
COLUMN_POSITIONS = {
    "DESCRIPTION": "G",  # Description skal være i kolonne G
    "BOM STRUCTURE": "F", # BOM Structure skal være i kolonne F
    "QTY": "J",         # QTY skal være i kolonne J
    "TOTAL QTY": "K"    # Total QTY skal være i kolonne K
}

EXCEL_FORMATTING = {
    "column_widths": {  # Bredde i pixels
        "A": 52,   # Item
        "B": 152,  # Part Number
        "C": 47,   # Rev
        "D": 111,  # Description 1
        "E": 93,   # Description 2
        "F": 115,  # BOM Structure
        "G": 423,  # Description
        "H": 135,  # Material
        "I": 173,  # Standard/PED
        "J": 48,   # QTY
        "K": 82,   # Total QTY
        "L": 39,   # Weight
        "M": 39,   # Surface Area
        "N": 39,   # Volume
        "O": 200,  # Comment
        "P": 94,   # Drawings
    },
    "row_heights": {
        "header": 20,  # Første række (header)
        "data": 91,    # Alle andre rækker
    }
}

# Excel formatering instruktioner
"""
Excel ark skal formateres på følgende måde:

Kolonnebredder (i pixels):
- A (Item): 52
- B (Part Number): 152
- C (Rev): 47
- D (Description 1): 111
- E (Description 2): 93
- F (BOM Structure): 115
- G (Description): 423
- H (Material): 135
- I (Standard/PED): 173
- J (QTY): 48
- K (Total QTY): 82
- L (Weight): 39
- M (Surface Area): 39
- N (Volume): 39
- O (Comment): 200
- P (Drawings): 94

Rækkehøjder:
- Første række (header): 20 pixels
- Alle andre rækker: 91 pixels

Andre formateringskrav:
- Første række skal have filter
- Første række skal være frossen (freeze pane)
- Rækker skal grupperes efter parent items
"""

# # Visual Basic Piping funktion reference
# VBA_PIPING_FUNCTION = 
# Sub subGroupTest()
#     Dim sRng As Range, eRng As Range
#     Dim groupMap() As Variant
#     Dim subGrp As Integer, i As Integer, j As Integer
#     Dim startRow As Range, lastRow As Range
#     Dim startGrp As Range, lastGrp As Range

#     ReDim groupMap(1 To 2, 1 To 1)
#     subGrp = 0
#     i = 0
#     Set startRow = Range("A1")

#     ' Create a map of the groups with their cell addresses and an index of the lowest subgrouping
#     Do While (startRow.Offset(i).Value <> "")
#         groupMap(1, i + 1) = startRow.Offset(i).Address
#         groupMap(2, i + 1) = UBound(Split(startRow.Offset(i).Value, "."))
#         If subGrp < groupMap(2, i + 1) Then subGrp = groupMap(2, i + 1)
#         ReDim Preserve groupMap(1 To 2, 1 To (i + 2))

#         Set lastRow = Range(groupMap(1, i + 1))
#         i = i + 1
#     Loop

#     ' Destroy already existing groups
#     On Error Resume Next
#     For k = 1 To 10
#         Rows(startRow.Row & ":" & lastRow.Row).EntireRow.Ungroup
#     Next k
#     On Error GoTo 0

#     ' Create the groups by levels in descending order
#     Do While (subGrp > 0)
#         For j = LBound(groupMap, 2) To UBound(groupMap, 2)
#             If groupMap(2, j) >= CStr(subGrp) Then
#                 If startGrp Is Nothing Then
#                     Set startGrp = Range(groupMap(1, j))
#                 End If
#                 Set lastGrp = Range(groupMap(1, j))
#             Else
#                 If Not startGrp Is Nothing And Not lastGrp Is Nothing Then 
#                     Range(startGrp, lastGrp).EntireRow.Group
#                 End If
#                 If Not startGrp Is Nothing Then Set startGrp = Nothing
#                 If Not lastGrp Is Nothing Then Set lastGrp = Nothing
#             End If
#         Next j
#         subGrp = subGrp - 1
#     Loop
# End Sub
# """ 

# Detaljerede instruktioner for hver step
STEP_INSTRUCTIONS = {
    "Indskyder linje i række 2": """
    1. Indsæt ny række 2
    2. Sæt værdier:
       - Item: "0"
       - Part Number: [projekt-nummer]-[sub-projekt]-A01
       - BOM Structure: "Inseparable"
       - Description: "Area Layout Drawing"
       - QTY: "1"
    """,
    
    "Fjerner Inseparable children": """
    1. Find alle rækker markeret som "Inseparable"
    2. Find og fjern alle children (rækker der starter med parent item number + ".")
    """,
    
    "Scanner for piping items": """
    1. Scan ALLE part numbers mod piping_categories.txt
    2. Special håndtering af 630-numre:
       - Hvis BG kode findes: Kategoriser som "Biogas Piping"
       - Ellers: Kategoriser som "Biomass Piping"
    3. For andre numre:
       - Tjek for BM prefix først (Biomass Piping)
       - Tjek derefter andre piping koder
    4. Gem hele gruppen (parent + children) under samme piping kategori
    """,
    
    "Scanner andre part numbers": """
    1. KUN for part numbers der IKKE er kategoriseret som piping
    2. For 0000-numre: Brug categories.txt
    3. For andre numre: Brug mappens navn
    4. Hvis ingen match: Kategoriser som "Other Items"
    """
} 