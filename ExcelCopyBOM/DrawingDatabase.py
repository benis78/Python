"""Database håndtering for tegninger"""

import sqlite3
from pathlib import Path
from typing import List, Dict
import config

class DrawingDatabase:
    """Klasse til at håndtere tegningsdatabase"""
    def __init__(self):
        self.db_path = Path(config.DRAWING_DB_PATH)
        
    def find_drawings(self, part_number: str) -> List[Dict[str, str]]:
        """Find tegninger for et part number"""
        if not part_number or not isinstance(part_number, str):
            return []
            
        try:
            # Tilslut til databasen
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Søg efter tegninger
            cursor.execute(
                "SELECT path, name FROM files WHERE name LIKE ?",
                (f"{part_number}%",)
            )
            
            # Konverter resultater til liste af dictionaries
            drawings = [
                {'filepath': row[0], 'filename': row[1]}
                for row in cursor.fetchall()
            ]
            
            return drawings
            
        except Exception as e:
            print(f"Fejl ved søgning i tegningsdatabase: {str(e)}")
            return []
            
        finally:
            if conn:
                conn.close() 