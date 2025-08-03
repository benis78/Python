"""
Test af ExcelOutlineHandler klassen
"""
import unittest
import pandas as pd
import queue
import tempfile
from pathlib import Path
from ExcelCopyBOM.threaded_excel_handler import ExcelOutlineHandler, ProcessingResult

class TestExcelOutlineHandler(unittest.TestCase):
    def setUp(self):
        """Setup test data og opret test Excel fil"""
        self.test_queue = queue.Queue()
        self.temp_dir = Path(tempfile.gettempdir()) / "ExcelCopyBOM_Test"
        self.temp_dir.mkdir(exist_ok=True)
        self.test_file = self.temp_dir / "test_outline.xlsx"
        
        # Opret test DataFrame
        self.test_df = pd.DataFrame({
            'Item': range(1, 8),
            'Part Number': [
                '4003-02.1-A01',  # Level 1
                '1234-56.1',      # Level 2
                '1234-57.1',      # Level 3
                '1234-58.1',      # Level 3
                '1235-56.1',      # Level 2
                '1235-57.1',      # Level 3
                '4003-02.2'       # Level 1
            ],
            'BOM Structure': [
                '1',        # Level 1
                '1.1',      # Level 2
                '1.1.1',    # Level 3
                '1.1.2',    # Level 3
                '1.2',      # Level 2
                '1.2.1',    # Level 3
                '2'         # Level 1
            ],
            'Description': [f'Test {i}' for i in range(1, 8)],
            'QTY': [1] * 7
        })
        
        # Gem test DataFrame til Excel
        self.test_df.to_excel(self.test_file, index=False)
        
        # Initialiser handler
        self.handler = ExcelOutlineHandler(self.test_file, self.test_queue)
        
    def tearDown(self):
        """Cleanup efter tests"""
        try:
            if self.test_file.exists():
                self.test_file.unlink()
        except Exception as e:
            print(f"Kunne ikke slette test fil: {str(e)}")
            
    def test_group_rows(self):
        """Test gruppering af rækker"""
        try:
            # Initialiser Excel
            self.handler._init_excel()
            
            # Test gruppering
            self.handler._group_rows(self.test_df)
            
            # Verify at filen stadig eksisterer og er blevet opdateret
            self.assertTrue(self.test_file.exists())
            self.assertGreater(self.test_file.stat().st_size, 0)
            
        finally:
            self.handler._cleanup_excel()
            
    def test_run_process(self):
        """Test hele outline processen"""
        # Put test data i queue
        self.test_queue.put(ProcessingResult(
            success=True,
            data={'dataframe': self.test_df}
        ))
        
        # Kør handler
        self.handler.run()
        
        # Check resultat
        self.assertTrue(self.handler.result.success)
        
    def test_invalid_input(self):
        """Test håndtering af ugyldig input"""
        # Test med manglende BOM Structure kolonne
        invalid_df = pd.DataFrame({
            'Part Number': ['Test1', 'Test2'],
            'Description': ['Desc1', 'Desc2']
        })
        
        invalid_df.to_excel(self.test_file, index=False)
        
        try:
            self.handler._init_excel()
            
            with self.assertRaises(Exception):
                self.handler._group_rows(invalid_df)
                
        finally:
            self.handler._cleanup_excel()
            
if __name__ == '__main__':
    unittest.main() 