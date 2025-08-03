import os
import sqlite3
import re

# Database-fil
sqlite_db = r"\\192.168.170.18\Drawings\file_index.db"

def search_files(search_term, file_type=None):
    """
    Søger i databasen efter filer der matcher søgetermerne
    :param search_term: Søgetekst (kan være partial match)
    :param file_type: Filtype at filtrere efter (f.eks. '.pdf')
    :return: Liste af fundne filer
    """
    try:
        conn = sqlite3.connect(sqlite_db)
        cursor = conn.cursor()
        
        search_pattern = f"%{search_term}%"
        
        if file_type:
            cursor.execute("""
                SELECT path, filename, size, modified_time 
                FROM files 
                WHERE (filename LIKE ? OR path LIKE ?) 
                AND file_type = ?
                ORDER BY modified_time DESC
            """, (search_pattern, search_pattern, file_type))
        else:
            cursor.execute("""
                SELECT path, filename, size, modified_time 
                FROM files 
                WHERE filename LIKE ? OR path LIKE ?
                ORDER BY modified_time DESC
            """, (search_pattern, search_pattern))
            
        results = cursor.fetchall()
        conn.close()
        
        # Formatér resultater
        formatted_results = []
        for row in results:
            formatted_results.append({
                'path': row[0],
                'filename': row[1],
                'size': row[2],
                'modified_time': row[3]
            })
            
        return formatted_results
    except Exception as e:
        print(f"Fejl ved søgning: {e}")
        return []

def main():
    """Hovedfunktion der demonstrerer søgning i databasen"""
    print("Filsøgning i SQLite database")
    print("-----------------------")
    
    if not os.path.exists(sqlite_db):
        print(f"Kan ikke finde databasen: {sqlite_db}")
        return
        
    while True:
        try:
            print("\nVælg en handling:")
            print("1. Søg efter filer")
            print("2. Vis alle PDF filer")
            print("3. Afslut")
            
            valg = input("\nIndtast dit valg (1-3): ").strip()
            
            if valg == "1":
                søgeord = input("Indtast søgeord: ").strip()
                filtype = input("Indtast filtype (f.eks. .pdf) eller tryk ENTER for alle: ").strip()
                
                if not filtype:
                    filtype = None
                    
                resultater = search_files(søgeord, filtype)
                
                if resultater:
                    print(f"\nFandt {len(resultater)} resultater:")
                    for res in resultater:
                        print(f"- {res['filename']} ({res['size']} bytes)")
                        print(f"  Sti: {res['path']}")
                else:
                    print("Ingen resultater fundet")
                    
            elif valg == "2":
                resultater = search_files("", ".pdf")
                if resultater:
                    print(f"\nFandt {len(resultater)} PDF filer:")
                    for res in resultater:
                        print(f"- {res['filename']} ({res['size']} bytes)")
                else:
                    print("Ingen PDF filer fundet")
                    
            elif valg == "3":
                print("Afslutter program...")
                break
                
            else:
                print("Ugyldigt valg, prøv igen")
                
        except Exception as e:
            print(f"Der opstod en fejl: {e}")
            
if __name__ == "__main__":
    main() 