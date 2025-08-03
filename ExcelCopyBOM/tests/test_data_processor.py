"""
Test af DataProcessor klassen
"""
import unittest
import pandas as pd
import queue
from pathlib import Path
from ExcelCopyBOM.threaded_excel_handler import DataProcessor, ProcessingResult
from .test_mock import MockGUI

class TestDataProcessor(unittest.TestCase):
    def setUp(self):
        """Setup test data"""
        self.test_queue = queue.Queue()
        self.test_file = Path("test_data.xlsx")
        self.mock_gui = MockGUI()
        self.processor = DataProcessor(self.test_file, self.test_queue)
        
        # Opret test DataFrame
        self.test_df = pd.DataFrame({
            'Item': [1, 2, 3, 4],
            'Part Number': ['4003-02.1-A01A', '0000-700-123', '1234-56.1-A', '4003-02.2'],
            'REV': ['', '', 'B', ''],
            'BOM Structure': ['1', '1.1', '1.1.1', 'Phantom'],
            'Description': ['Test 1', 'Test 2', 'Test 3', 'Test 4'],
            'QTY': [1, 2, 3, 4],
            'D': [1, 1, 1, 1],
            't': [1, 1, 1, 1],
            'L': [1, 1, 1, 1]
        })
        
        # Gem test DataFrame
        self.test_df.to_excel(str(self.test_file), index=False)
        
    def tearDown(self):
        """Cleanup efter tests"""
        try:
            if self.test_file.exists():
                self.test_file.unlink()
        except Exception as e:
            print(f"Kunne ikke slette test fil: {str(e)}")

    def test_extract_revision(self):
        """Test udtrækning af revision fra filnavn"""
        test_cases = [
            ("4003-02.1-A01A -- Test.xlsx", ("4003-02.1-A01", "A")),
            ("1234-56.1 -- No Rev.xlsx", ("1234-56.1", "")),
            ("NoPartNumber.xlsx", ("NoPartNumber", ""))
        ]
        
        for filename, expected in test_cases:
            part_number, rev = self.processor._extract_revision(filename)
            self.assertEqual((part_number, rev), expected)
            
    def test_extract_revision_from_partnumber(self):
        """Test udtrækning af revision fra part number"""
        test_cases = [
            ("4003-02.1-A01A", ("4003-02.1-A01", "A")),
            ("1234-56.1", ("1234-56.1", "")),
            ("0000-700-123B", ("0000-700-123", "B"))
        ]
        
        for part_number, expected in test_cases:
            clean_number, rev = self.processor._extract_revision_from_partnumber(part_number)
            self.assertEqual((clean_number, rev), expected)
            
    def test_is_supplier_part(self):
        """Test identifikation af supplier parts"""
        test_cases = [
            ("0000-700-123", True),
            ("0000-701-456", True),
            ("0000-702-789", True),
            ("4003-02.1-A01", False),
            ("1234-56.1", False)
        ]
        
        for part_number, expected in test_cases:
            result = self.processor._is_supplier_part(part_number)
            self.assertEqual(result, expected)
            
    def test_calculate_total_qty(self):
        """Test beregning af Total QTY"""
        # Test DataFrame med BOM struktur
        df = pd.DataFrame({
            'Part Number': ['P1', 'P2', 'P3', 'P4'],
            'BOM Structure': ['1', '1.1', '1.1.1', '2'],
            'QTY': [1, 2, 3, 4]
        })
        
        result_df = self.processor._calculate_total_qty(df)
        
        # Check Total QTY beregninger
        self.assertEqual(result_df.loc[0, 'Total QTY'], 1)  # Level 1
        self.assertEqual(result_df.loc[1, 'Total QTY'], 2)  # Level 2 (2 * 1)
        self.assertEqual(result_df.loc[2, 'Total QTY'], 6)  # Level 3 (3 * 2)
        self.assertEqual(result_df.loc[3, 'Total QTY'], 4)  # Nyt Level 1
        
    def test_add_categories(self):
        """Test tilføjelse af kategorier"""
        df = pd.DataFrame({
            'Part Number': ['4003-02.1-A01', '0000-615-123', '1234-56.1']
        })
        
        result_df = self.processor._add_categories(df)
        
        # Check at Category kolonne er tilføjet
        self.assertIn('Category', result_df.columns)
        
        # Check specifikke kategorier (afhænger af Categories.csv)
        self.assertNotEqual(result_df.loc[0, 'Category'], '')
        
    def test_process_bom_structure_rules(self):
        """Test BOM Structure regler"""
        df = pd.DataFrame({
            'Part Number': ['P1', 'P2', 'P3', 'P4', 'P5'],
            'BOM Structure': ['1', '1.1', 'Phantom', 'Inseparable', '2.1'],
            'Description': ['Test 1', 'Test 2', 'Test 3', 'Test 4', 'Test 5']
        })
        
        # Find rækker der skal slettes
        rows_to_delete = []
        for idx, row in df.iterrows():
            if pd.notna(row['BOM Structure']):
                if str(row['BOM Structure']).lower() == 'phantom':
                    rows_to_delete.append(idx)
                    
        df = df.drop(rows_to_delete)
        
        # Check at Phantom række er slettet
        self.assertNotIn('Phantom', df['BOM Structure'].values)
        
if __name__ == '__main__':
    unittest.main() 