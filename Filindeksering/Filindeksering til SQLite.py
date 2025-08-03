import os
import sqlite3
import re

# Database-fil
sqlite_db = "file_index.db"

# Funktion til at udtrække grundlæggende filnavn og revision
def extract_filename_info(filename):
    match = re.match(r"^(.*?)(-[A-Z]?)?(\..+)$", filename)  # Matcher 4003-02.1-A01 og revision (A, B, osv.)
    if match:
        base_name = match.group(1)  # F.eks. "4003-02.1-A01"
        revision = match.group(2) if match.group(2) else "-"  # Hvis ingen revision, brug "-"
        return base_name, revision
    return filename, "-"

# Funktion til at indeksere netværksdrevets PDF-filer
def create_sqlite_index(root_dir, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Opret tabel med metadata
    cursor.execute("DROP TABLE IF EXISTS files")
    cursor.execute("""
        CREATE TABLE files (
            path TEXT PRIMARY KEY,
            filename TEXT,
            base_name TEXT,
            revision TEXT,
            size INTEGER,
            modified_time REAL
        )
    """)

    for dirpath, _, filenames in os.walk(root_dir):
        for file in filenames:
            if file.endswith(".pdf"):
                full_path = os.path.join(dirpath, file)
                size = os.path.getsize(full_path)
                modified_time = os.path.getmtime(full_path)
                base_name, revision = extract_filename_info(file)
                
                cursor.execute("INSERT INTO files VALUES (?, ?, ?, ?, ?, ?)",
                               (full_path, file, base_name, revision, size, modified_time))
    
    conn.commit()
    conn.close()

# Indekser hele netværksdrevet (kun nødvendigt én gang)
network_drive = r'\\192.168.170.18\drawings'
create_sqlite_index(network_drive, sqlite_db)
