"""
Database handling with caching for ExcelCopyBOM
"""
import sqlite3
import threading
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor
import subprocess
from . import config

class DatabaseCache:
    def __init__(self):
        self._cache: Dict[str, Tuple[float, List[Dict]]] = {}
        self._lock = threading.Lock()
        self._executor = ThreadPoolExecutor(max_workers=1)
        
    def get(self, key: str) -> Optional[List[Dict]]:
        """Hent værdi fra cache hvis den findes og ikke er for gammel"""
        with self._lock:
            if key in self._cache:
                timestamp, value = self._cache[key]
                if time.time() - timestamp < config.DB_CACHE_TIMEOUT:
                    return value
                else:
                    del self._cache[key]
            return None
            
    def set(self, key: str, value: List[Dict]):
        """Gem værdi i cache med timestamp"""
        with self._lock:
            if len(self._cache) >= config.DB_CACHE_SIZE:
                # Fjern ældste entry hvis cachen er fuld
                oldest_key = min(self._cache.keys(), 
                               key=lambda k: self._cache[k][0])
                del self._cache[oldest_key]
            self._cache[key] = (time.time(), value)

class DrawingDatabase:
    def __init__(self):
        self._cache = DatabaseCache()
        self._db_path = Path(config.DB_PATH)
        self._executor = ThreadPoolExecutor(max_workers=1)
        
    def update_index(self) -> bool:
        """Kør file_indexer.exe for at opdatere databasen"""
        try:
            result = subprocess.run([config.FILE_INDEXER], 
                                  capture_output=True, 
                                  text=True)
            return result.returncode == 0
        except Exception as e:
            print(f"Error updating index: {e}")
            return False
            
    def _execute_query(self, query: str, params: tuple) -> List[Dict]:
        """Udfør SQL query i separat thread"""
        conn = sqlite3.connect(str(self._db_path))
        try:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
        finally:
            conn.close()
            
    def find_drawings(self, part_number: str) -> List[Dict]:
        """Find tegninger for et part number"""
        cache_key = f"drawings_{part_number}"
        result = self._cache.get(cache_key)
        if result is not None:
            return result
            
        # Søg efter både PDF og DWG filer
        query = """
            SELECT path, filename, file_type, modified_time
            FROM files
            WHERE filename LIKE ? 
            AND file_type IN ('.pdf', '.dwg')
            ORDER BY modified_time DESC
        """
        
        future = self._executor.submit(
            self._execute_query, 
            query, 
            (f"{part_number}%",)
        )
        result = future.result()
        
        self._cache.set(cache_key, result)
        return result
        
    def find_latest_revision(self, part_number: str) -> Optional[str]:
        """Find seneste revision for et part number"""
        drawings = self.find_drawings(part_number)
        if not drawings:
            return None
            
        # Find seneste revision fra filnavnet
        latest_rev = None
        for drawing in drawings:
            filename = Path(drawing['filename']).stem
            parts = filename.split('-')
            if len(parts) > 1:
                rev = parts[-1]
                if not latest_rev or rev > latest_rev:
                    latest_rev = rev
                    
        return latest_rev
        
    def get_drawing_status(self, part_number: str) -> str:
        """
        Returner tegningsstatus:
        - DWG_PDF: Begge filer findes med samme revision
        - DWG~PDF: Begge filer findes med forskellige revisioner
        - DWG: Kun DWG findes
        - PDF: Kun PDF findes
        - "": Ingen filer findes
        """
        drawings = self.find_drawings(part_number)
        if not drawings:
            return ""
            
        has_pdf = any(d['file_type'].lower() == '.pdf' for d in drawings)
        has_dwg = any(d['file_type'].lower() == '.dwg' for d in drawings)
        
        if has_pdf and has_dwg:
            # Tjek om revisionerne er ens
            pdf_rev = next(d['filename'].split('-')[-1] for d in drawings 
                         if d['file_type'].lower() == '.pdf')
            dwg_rev = next(d['filename'].split('-')[-1] for d in drawings 
                         if d['file_type'].lower() == '.dwg')
            
            return "DWG_PDF" if pdf_rev == dwg_rev else "DWG~PDF"
        elif has_dwg:
            return "DWG"
        elif has_pdf:
            return "PDF"
        return "" 