"""
Test af FileProcessor klassen
"""
import unittest
import tempfile
import shutil
from pathlib import Path
from ExcelCopyBOM.threaded_excel_handler import FileProcessor

class TestFileProcessor(unittest.TestCase):
    def setUp(self):
        """Setup test data og mapper"""
        # Opret temp directories
        self.temp_dir = Path(tempfile.gettempdir()) / "ExcelCopyBOM_Test"
        self.source_dir = self.temp_dir / "source"
        self.target_dir = self.temp_dir / "target"
        
        self.source_dir.mkdir(parents=True, exist_ok=True)
        self.target_dir.mkdir(parents=True, exist_ok=True)
        
        # Opret test filer
        self.test_files = [
            "4003-02.1-A01-- Test1.dwg",
            "4003-02.1-A01-- Test1.pdf",
            "1234-56.1-- Test2.dwg",
            "0000-615-123-- Test3.pdf"
        ]
        
        for filename in self.test_files:
            test_file = self.source_dir / filename
            test_file.write_text("Test content")
            
        # Definer kategori mapping
        self.categories_map = {
            "4003-02.1-A01": "Tank Drawings",
            "1234-56.1": "Equipment Drawings",
            "0000-615-123": "Primary Digester"
        }
        
        # Initialiser processor
        self.source_files = [self.source_dir / f for f in self.test_files]
        self.processor = FileProcessor(self.source_files, self.target_dir, self.categories_map)
        
    def tearDown(self):
        """Cleanup efter tests"""
        try:
            shutil.rmtree(self.temp_dir)
        except Exception as e:
            print(f"Kunne ikke slette test mapper: {str(e)}")
            
    def test_create_category_folders(self):
        """Test oprettelse af kategori-mapper"""
        category_paths = self.processor._create_category_folders()
        
        # Check at alle kategori-mapper er oprettet
        for category in self.categories_map.values():
            category_path = self.target_dir / category
            self.assertTrue(category_path.exists())
            self.assertTrue(category_path.is_dir())
            
        # Check at mapping er korrekt
        self.assertEqual(len(category_paths), len(set(self.categories_map.values())))
        
    def test_copy_file(self):
        """Test kopiering af enkelt fil"""
        source_file = self.source_files[0]
        target_file = self.target_dir / source_file.name
        
        # Test succesfuld kopiering
        success = self.processor._copy_file(source_file, target_file)
        self.assertTrue(success)
        self.assertTrue(target_file.exists())
        self.assertEqual(target_file.stat().st_size, source_file.stat().st_size)
        
        # Test kopiering af ikke-eksisterende fil
        invalid_source = self.source_dir / "non_existent.txt"
        invalid_target = self.target_dir / "non_existent.txt"
        success = self.processor._copy_file(invalid_source, invalid_target)
        self.assertFalse(success)
        
    def test_run_process(self):
        """Test hele fil-processerings processen"""
        # Kør processor
        self.processor.run()
        
        # Check resultat
        self.assertTrue(self.processor.result.success)
        
        # Check at filer er kopieret til rigtige mapper
        for source_file in self.source_files:
            part_number = source_file.stem.split("--")[0].strip()
            if part_number in self.categories_map:
                category = self.categories_map[part_number]
                target_path = self.target_dir / category / source_file.name
                self.assertTrue(target_path.exists())
                self.assertEqual(target_path.stat().st_size, source_file.stat().st_size)
                
    def test_invalid_part_numbers(self):
        """Test håndtering af ukendte part numbers"""
        # Opret fil med ukendt part number
        invalid_file = self.source_dir / "9999-99.9-- Unknown.dwg"
        invalid_file.write_text("Test content")
        
        # Tilføj til source files
        self.processor.source_files.append(invalid_file)
        
        # Kør processor
        self.processor.run()
        
        # Check at filen ikke er kopieret
        for category_dir in self.target_dir.iterdir():
            if category_dir.is_dir():
                self.assertFalse((category_dir / invalid_file.name).exists())
                
if __name__ == '__main__':
    unittest.main() 