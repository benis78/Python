"""Konfiguration for ExcelCopyBOM"""

# Kolonner der skal kopieres
REQUIRED_COLUMNS = [
    'Item',
    'Part Number',
    'REV',
    'Description 1',
    'Thumbnail',
    'BOM Structure',
    'Description',
    'Material',
    'Title',
    'QTY',
    'Total QTY',
    'Diameter',
    'Thickness',
    'Length',
    'Keywords',
    'Type',
    'Category',
    'Drawings'
]

# Leverandør prefixes
SUPPLIER_PREFIXES = [
    '0000-700',
    '0000-701',
    '0000-702'
]

# Kategorier
CATEGORIES = {
    '0000-1': 'Købt',
    '0000-2': 'Bearbejdet',
    '0000-3': 'Samling',
    '0000-4': 'Tegning',
    '0000-5': 'Støbt',
    '0000-6': 'Svejst',
    '0000-7': 'Leverandør'
}

# Kolonne bredder (i pixels)
# Note: Excel bruger en anden enhed for kolonnebredder, så vi konverterer fra pixels
# Cirka konvertering: pixels / 7.5 = Excel bredde enheder
COLUMN_WIDTHS = {
    'A': 52 / 7.5,  # Item
    'B': 152 / 7.5,  # Part Number
    'C': 47 / 7.5,  # Rev
    'D': 111 / 7.5,  # Description 1
    'E': 93 / 7.5,  # Thumbnail
    'F': 115 / 7.5,  # BOM Structure
    'G': 423 / 7.5,  # Description
    'H': 135 / 7.5,  # Material
    'I': 173 / 7.5,  # Title
    'J': 48 / 7.5,  # QTY
    'K': 82 / 7.5,  # Total QTY
    'L': 39 / 7.5,  # Diameter
    'M': 39 / 7.5,  # Thickness
    'N': 39 / 7.5,  # Length
    'O': 200 / 7.5,  # Keywords
    'P': 180 / 7.5,  # Type
    'Q': 125 / 7.5,  # Category
    'R': 94 / 7.5,  # Drawings
}

# Excel formatering
EXCEL_FORMATTING = {
    'header_row_height': 26,  # pixels
    'data_row_height': 91,    # pixels
    'header_row_bold': True,
    'freeze_first_row': True,
    'auto_filter': True
}

# Tegnings database sti
DRAWING_DB_PATH = r"\\192.168.170.18\Drawings\file_index.db"
DRAWING_INDEXER_PATH = r"\\192.168.170.18\Drawings\file_indexer.exe"

# Compare sheet formatering
COMPARE_COLORS = {
    'new_row': (198, 239, 206),      # Lysegrøn
    'changed_cell': (255, 235, 156),  # Lysgul
    'deleted_row': (255, 199, 206)    # Lysrød
}

# Drawing status værdier
DRAWING_STATUS = {
    'both_match': 'DWG_PDF',    # Både DWG og PDF findes og revision matcher
    'both_mismatch': 'DWG!PDF', # Både DWG og PDF findes men revision matcher ikke
    'pdf_only': 'PDF',          # Kun PDF findes
    'dwg_only': 'DWG',          # Kun DWG findes
    'none': ''                  # Ingen tegninger fundet
} 