"""
Konfigurationsfil for ExcelCopyBOM
"""
import os
from pathlib import Path

# Base paths
BASE_DIR = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DRAWING_DIR = Path("//SERVER/Drawings")  # Sti til tegninger på server

# Database indstillinger
DB_FILE = BASE_DIR / "drawing_index.db"

# Excel indstillinger
EXCEL_TEMPLATE = BASE_DIR / "template.xlsx"
DEFAULT_SHEET = "BOM"

# GUI indstillinger
WINDOW_TITLE = "Excel Copy BOM"
WINDOW_SIZE = "800x600"

# Temporary directory for image handling
TEMP_DIR = Path(Path.home() / "AppData" / "Local" / "Temp" / "ExcelCopyBOM")
TEMP_DIR.mkdir(parents=True, exist_ok=True)

# Required Excel columns
REQUIRED_COLUMNS = [
    "Item",
    "Qty",
    "Part Number",
    "Description",
    "REV",
    "Category"
]

# GUI settings
PROGRESS_UPDATE_MS = 100  # Milliseconds between progress updates

# Cache settings
DB_CACHE_SIZE = 1000  # Number of entries to cache
DB_CACHE_TIMEOUT = 300  # 5 minutes

# Logging
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'

# Excel formatting
HEADER_ROW_HEIGHT = 26
DATA_ROW_HEIGHT = 91
COLUMN_WIDTHS = {
    "Item": 10,
    "Qty": 10,
    "Part Number": 20,
    "Description": 40,
    "REV": 10,
    "Category": 25,
    "Drawing Status": 15
}

# Supplier part numbers to exclude
SUPPLIER_PATTERNS = [
    "0000-700",
    "0000-701",
    "0000-702"
]

# Category settings
DEFAULT_CATEGORY = "Other Parts"
SPECIAL_CATEGORIES = {
    "Piping": ["Biomass Piping", "Process Piping"]
}

# Sheet names
PARTLIST_SHEET = "Partlist"
COMPARE_SHEET = "Compare" 