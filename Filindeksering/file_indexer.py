"""
Filindeksering til SQLite
-------------------------
Dette modul indekserer filer på et netværksdrev og gemmer deres metadata i en SQLite-database.
"""

import os
import sqlite3
import logging
import time
from datetime import datetime

# Konfiguration
NETVAERKSDREV = r'\\192.168.170.18\Drawings'
SQLITE_DB = os.path.join(NETVAERKSDREV, "file_index.db")
LOG_FILE = os.path.join(NETVAERKSDREV, "filindeksering.log")

# Opsæt logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('FilIndeksering')

def verify_database():
    """Verificerer at databasen eksisterer og har den korrekte struktur"""
    try:
        print(f"Tjekker database: {SQLITE_DB}")
        if not os.path.exists(SQLITE_DB):
            print("Database fil findes ikke endnu")
            logger.info("Database fil findes ikke endnu")
            return False
        else:
            print("Database fil fundet")
            
        try:
            conn = sqlite3.connect(SQLITE_DB)
            cursor = conn.cursor()
            
            # Tjek om de nødvendige tabeller findes
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name IN ('files', 'metadata')
            """)
            tables = cursor.fetchall()
            tables = [table[0] for table in tables]
            
            print(f"Fundne tabeller: {tables}")
            
            if 'files' not in tables or 'metadata' not in tables:
                print("Manglende tabeller i databasen")
                logger.info("Manglende tabeller i databasen")
                conn.close()
                return False
            
            # Tjek metadata tabel for last_scan_time
            cursor.execute("SELECT key, value FROM metadata")
            metadata = dict(cursor.fetchall())
            print(f"Metadata indhold: {metadata}")
            
            if 'last_scan_time' not in metadata:
                print("Ingen sidste scanningstidspunkt fundet")
                logger.info("Ingen sidste scanningstidspunkt fundet")
                conn.close()
                return False
            
            conn.close()
            print("Database verificeret korrekt")
            logger.info("Database verificeret korrekt")
            return True
            
        except sqlite3.Error as e:
            print(f"SQLite fejl: {e}")
            logger.error(f"SQLite fejl: {e}")
            return False
            
    except Exception as e:
        print(f"Fejl ved verificering af database: {e}")
        logger.error(f"Fejl ved verificering af database: {e}")
        return False

def create_database():
    """Opretter eller opdaterer SQLite-databasen med den korrekte tabelstruktur"""
    try:
        # Tjek først om databasen allerede eksisterer og er valid
        if verify_database():
            logger.info("Eksisterende database fundet og verificeret")
            return True
            
        logger.info("Opretter/opdaterer database struktur")
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        
        # Opret tabeller hvis de ikke findes
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                path TEXT UNIQUE,
                filename TEXT,
                size INTEGER,
                modified_time REAL,
                creation_time REAL,
                file_type TEXT
            )
        """)
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS metadata (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        
        # Opret indekser for hurtigere søgning
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_filename ON files(filename)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_file_type ON files(file_type)")
        
        # Hvis der ikke er et sidste scanningstidspunkt, sæt det til 0
        cursor.execute("""
            INSERT OR IGNORE INTO metadata (key, value)
            VALUES ('last_scan_time', '0')
        """)
        
        conn.commit()
        conn.close()
        logger.info("Database struktur oprettet/opdateret korrekt")
        return True
    except Exception as e:
        logger.error(f"Fejl ved oprettelse/opdatering af database: {e}")
        return False

def get_last_scan_time():
    """Henter tidspunktet for sidste scanning"""
    try:
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM metadata WHERE key = 'last_scan_time'")
        result = cursor.fetchone()
        conn.close()
        
        if result:
            return float(result[0])
        return 0
    except Exception as e:
        logger.error(f"Fejl ved hentning af sidste scanningstidspunkt: {e}")
        return 0

def update_last_scan_time():
    """Opdaterer tidspunktet for sidste scanning"""
    try:
        current_time = time.time()
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)",
                      ('last_scan_time', str(current_time)))
        conn.commit()
        conn.close()
        return current_time
    except Exception as e:
        logger.error(f"Fejl ved opdatering af sidste scanningstidspunkt: {e}")
        return None

def scan_directory(root_dir):
    """Scanner hele mappen og dens undermapper og indekserer filerne"""
    start_time = time.time()
    file_count = 0
    
    try:
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        
        # Scan alle mapper og undermapper
        for dirpath, _, filenames in os.walk(root_dir):
            for filename in filenames:
                try:
                    full_path = os.path.join(dirpath, filename)
                    
                    # Få filoplysninger
                    file_stats = os.stat(full_path)
                    size = file_stats.st_size
                    modified_time = file_stats.st_mtime
                    creation_time = file_stats.st_ctime
                    file_type = os.path.splitext(filename)[1].lower()
                    
                    # Indsæt eller opdater i databasen
                    cursor.execute("""
                        INSERT OR REPLACE INTO files 
                        (path, filename, size, modified_time, creation_time, file_type)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (full_path, filename, size, modified_time, creation_time, file_type))
                    
                    file_count += 1
                    
                    # Commit for hver 1000 filer for at undgå for stor transaktion
                    if file_count % 1000 == 0:
                        conn.commit()
                        logger.info(f"Indekseret {file_count} filer...")
                        
                except Exception as e:
                    logger.error(f"Fejl ved indeksering af {full_path}: {e}")
        
        # Opdater sidste scanningstidspunkt
        current_time = time.time()
        cursor.execute("UPDATE metadata SET value = ? WHERE key = 'last_scan_time'", (str(current_time),))
        
        conn.commit()
        conn.close()
        
        elapsed_time = time.time() - start_time
        logger.info(f"Scanning afsluttet. Indekseret {file_count} filer på {elapsed_time:.2f} sekunder")
        return True
    except Exception as e:
        logger.error(f"Fejl under scanning af mappe {root_dir}: {e}")
        if 'conn' in locals():
            conn.close()
        return False

def scan_for_changes(root_dir):
    """Scanner efter nye eller ændrede filer siden sidste scanning"""
    start_time = time.time()
    file_count = 0
    last_scan_time = get_last_scan_time()
    
    try:
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        
        # Scan alle mapper og undermapper
        for dirpath, _, filenames in os.walk(root_dir):
            for filename in filenames:
                try:
                    full_path = os.path.join(dirpath, filename)
                    
                    # Tjek filens ændringstidspunkt
                    file_stats = os.stat(full_path)
                    modified_time = file_stats.st_mtime
                    
                    # Hvis filen er nyere end sidste scanning
                    if modified_time > last_scan_time:
                        size = file_stats.st_size
                        creation_time = file_stats.st_ctime
                        file_type = os.path.splitext(filename)[1].lower()
                        
                        # Indsæt eller opdater i databasen
                        cursor.execute("""
                            INSERT OR REPLACE INTO files 
                            (path, filename, size, modified_time, creation_time, file_type)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (full_path, filename, size, modified_time, creation_time, file_type))
                        
                        file_count += 1
                        
                        # Commit for hver 1000 filer
                        if file_count % 1000 == 0:
                            conn.commit()
                            logger.info(f"Opdateret {file_count} filer...")
                        
                except Exception as e:
                    logger.error(f"Fejl ved indeksering af {full_path}: {e}")
        
        # Fjern slettede filer fra databasen
        cursor.execute("SELECT path FROM files")
        db_files = set(row[0] for row in cursor.fetchall())
        
        existing_files = set()
        for dirpath, _, filenames in os.walk(root_dir):
            for filename in filenames:
                existing_files.add(os.path.join(dirpath, filename))
        
        # Fjern filer der ikke længere eksisterer
        deleted_files = db_files - existing_files
        for deleted_file in deleted_files:
            cursor.execute("DELETE FROM files WHERE path = ?", (deleted_file,))
            logger.info(f"Fjernet slettet fil fra indeks: {deleted_file}")
        
        conn.commit()
        conn.close()
        
        # Opdater sidste scanningstidspunkt
        update_last_scan_time()
        
        elapsed_time = time.time() - start_time
        logger.info(f"Ændringer scannet. Opdateret {file_count} filer på {elapsed_time:.2f} sekunder")
        return True
    except Exception as e:
        logger.error(f"Fejl under scanning af ændringer i {root_dir}: {e}")
        if 'conn' in locals():
            conn.close()
        return False

def search_files(search_term, file_type=None, limit=100):
    """
    Søger i databasen efter filer der matcher søgetermerne
    :param search_term: Søgetekst (kan være partial match)
    :param file_type: Filtype at filtrere efter (f.eks. '.pdf')
    :param limit: Maksimalt antal resultater
    :return: Liste af fundne filer
    """
    try:
        conn = sqlite3.connect(SQLITE_DB)
        conn.row_factory = sqlite3.Row  # Gør resultater til dictionaries
        cursor = conn.cursor()
        
        search_pattern = f"%{search_term}%"
        
        if file_type:
            cursor.execute("""
                SELECT * FROM files 
                WHERE (filename LIKE ? OR path LIKE ?) AND file_type = ?
                ORDER BY modified_time DESC
                LIMIT ?
            """, (search_pattern, search_pattern, file_type, limit))
        else:
            cursor.execute("""
                SELECT * FROM files 
                WHERE filename LIKE ? OR path LIKE ?
                ORDER BY modified_time DESC
                LIMIT ?
            """, (search_pattern, search_pattern, limit))
            
        results = cursor.fetchall()
        conn.close()
        
        # Formatér resultater
        formatted_results = []
        for row in results:
            formatted_results.append({
                'filename': row['filename'],
                'path': row['path'],
                'size': row['size'],
                'modified_time': datetime.fromtimestamp(row['modified_time']).strftime('%Y-%m-%d %H:%M:%S'),
                'creation_time': datetime.fromtimestamp(row['creation_time']).strftime('%Y-%m-%d %H:%M:%S'),
                'file_type': row['file_type']
            })
            
        return formatted_results
    except Exception as e:
        logger.error(f"Fejl ved søgning: {e}")
        return []

def main():
    """Hovedfunktion der tjekker for ændringer i filerne"""
    print("Filindeksering til SQLite")
    print("-----------------------")
    
    try:
        print(f"Bruger database: {SQLITE_DB}")
        
        # Verificer/opret database struktur
        if not create_database():
            print("Kunne ikke oprette/verificere database, afslutter...")
            return
        
        # Tjek om dette er første kørsel
        last_scan_time = get_last_scan_time()
        if last_scan_time == 0:
            print("Ingen tidligere scanning fundet - starter fuld indeksering...")
            if not scan_directory(NETVAERKSDREV):
                print("Fejl under første indeksering")
                return
        else:
            last_scan_datetime = datetime.fromtimestamp(last_scan_time)
            print(f"Tidligere scanning fundet fra: {last_scan_datetime}")
            print("Scanner efter ændringer...")
            if not scan_for_changes(NETVAERKSDREV):
                print("Fejl under scanning af ændringer")
                return
            
        print("\nScanning gennemført!")
        print(f"Database gemt i: {SQLITE_DB}")
        print("\nDu kan nu søge i databasen med search_files() funktionen")
        print("Kør programmet igen for at tjekke for nye ændringer")
            
    except Exception as e:
        logger.error(f"Uventet fejl: {e}")
        print(f"Der opstod en fejl: {e}")

if __name__ == "__main__":
    main() 