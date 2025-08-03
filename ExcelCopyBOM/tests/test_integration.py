"""
Integration test af hele ExcelCopyBOM processen
"""
import unittest
import tempfile
import pandas as pd
from pathlib import Path
from ExcelCopyBOM.threaded_excel_handler import ThreadedExcelHandler
from ExcelCopyBOM.gui import ExcelCopyBOMGUI
from tests.test_mock import MockGUI

class TestExcelCopyBOMIntegration(unittest.TestCase):
    def setUp(self):
        """Setup test data og miljø"""
        # Opret temp directories
        self.temp_dir = Path(tempfile.gettempdir()) / "ExcelCopyBOM_Test"
        self.input_dir = self.temp_dir / "input"
        self.output_dir = self.temp_dir / "output"
        
        self.input_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Opret test Excel fil
        self.test_file = self.input_dir / "test_bom.xlsx"
        self.test_df = pd.DataFrame({
            'Item': range(1, 5),
            'Part Number': [
                '4003-02.1-A01A',  # Level 1
                '0000-700-123',    # Supplier part
                '1234-56.1-A',     # Level 2
                '4003-02.2'        # Level 1
            ],
            'REV': ['', '', 'B', ''],
            'BOM Structure': ['1', '1.1', '1.1.1', '2'],
            'Description': ['Test 1', 'Test 2', 'Test 3', 'Test 4'],
            'QTY': [1, 2, 3, 4],
            'D': [1] * 4,
            't': [1] * 4,
            'L': [1] * 4
        })
        
        self.test_df.to_excel(str(self.test_file), index=False)
        
        # Opret mock GUI
        self.mock_gui = MockGUI()
        self.mock_gui.bom_path.get = lambda: str(self.test_file)
        
    def tearDown(self):
        """Cleanup efter tests"""
        import shutil
        try:
            shutil.rmtree(self.temp_dir)
        except Exception as e:
            print(f"Kunne ikke slette test mapper: {str(e)}")
            
    def test_full_process(self):
        """Test hele processen fra start til slut"""
        # Initialiser handler
        handler = ThreadedExcelHandler(self.test_file, self.output_dir)
        
        # Start processing
        success = handler.process()
        
        # Check resultat
        self.assertTrue(success)
        
        # Check output fil
        output_file = self.output_dir / f"{self.test_file.stem}_Processed.xlsx"
        self.assertTrue(output_file.exists())
        
        # Load output og check indhold
        result_df = pd.read_excel(str(output_file))
        
        # Check at supplier parts er fjernet
        self.assertNotIn('0000-700-123', result_df['Part Number'].values)
        
        # Check at Total QTY er beregnet
        self.assertIn('Total QTY', result_df.columns)
        
        # Check at Category er tilføjet
        self.assertIn('Category', result_df.columns)
        
if __name__ == '__main__':
    unittest.main() 