import unittest
import pandas as pd
import queue
from pathlib import Path
from ExcelCopyBOM.threaded_excel_handler import DataProcessor, ProcessingResult

class TestDataProcessor(unittest.TestCase):
    def setUp(self):
        """Opsæt test miljø"""
        self.test_file = Path("tests/test_data/test_bom-A -- Test.xlsx")
        self.output_queue = queue.Queue()
        self.processor = DataProcessor(self.test_file, self.output_queue)
        
    def test_extract_revision(self):
        """Test _extract_revision metoden"""
        # Test normal case
        part_number, rev = self.processor._extract_revision("1234-56A -- Test")
        self.assertEqual(part_number, "1234-56")
        self.assertEqual(rev, "A")
        
        # Test uden revision
        part_number, rev = self.processor._extract_revision("1234-56 -- Test")
        self.assertEqual(part_number, "1234-56")
        self.assertEqual(rev, "")
        
        # Test uden beskrivelse
        part_number, rev = self.processor._extract_revision("1234-56A")
        self.assertEqual(part_number, "1234-56A")
        self.assertEqual(rev, "")
        
    def test_extract_revision_from_partnumber(self):
        """Test _extract_revision_from_partnumber metoden"""
        # Test normal case
        part_number, rev = self.processor._extract_revision_from_partnumber("1234-56-A01A")
        self.assertEqual(part_number, "1234-56-A01")
        self.assertEqual(rev, "A")
        
        # Test uden revision
        part_number, rev = self.processor._extract_revision_from_partnumber("1234-56-A01")
        self.assertEqual(part_number, "1234-56-A01")
        self.assertEqual(rev, "")
        
        # Test med None
        part_number, rev = self.processor._extract_revision_from_partnumber(None)
        self.assertEqual(part_number, "None")
        self.assertEqual(rev, "")
        
    def test_is_supplier_part(self):
        """Test _is_supplier_part metoden"""
        # Test supplier parts
        self.assertTrue(self.processor._is_supplier_part("0000-700-123"))
        self.assertTrue(self.processor._is_supplier_part("0000-701-456"))
        self.assertTrue(self.processor._is_supplier_part("0000-702-789"))
        
        # Test ikke-supplier parts
        self.assertFalse(self.processor._is_supplier_part("1234-56"))
        self.assertFalse(self.processor._is_supplier_part("0000-703-123"))
        self.assertFalse(self.processor._is_supplier_part(None))
        
    def test_calculate_total_qty(self):
        """Test _calculate_total_qty metoden"""
        # Opret test DataFrame
        data = {
            'Part Number': ['1234-01', '1234-02', '1234-03', '1234-04'],
            'BOM Structure': ['1', '1.1', '1.1.1', '2'],
            'QTY': [1, 2, 3, 1],
            'Total QTY': [1, 2, 6, 1]  # Forventede værdier
        }
        df = pd.DataFrame(data)
        
        # Test beregning
        result_df = self.processor._calculate_total_qty(df)
        pd.testing.assert_series_equal(
            result_df['Total QTY'],
            pd.Series([1.0, 2.0, 6.0, 1.0], name='Total QTY')
        )
        
    def test_process_bom_structure_rules(self):
        """Test _process_bom_structure_rules metoden"""
        # Opret test DataFrame
        data = {
            'Part Number': ['1234-01', '1234-02', '0000-301', '1234-03', '1234-04'],
            'BOM Structure': ['1', '1.1', '2', '2.1', 'Phantom'],
            'QTY': [1, 2, 1, 2, 1]
        }
        df = pd.DataFrame(data)
        
        # Test regler
        result_df = self.processor._process_bom_structure_rules(df)
        
        # Verificer at Phantom og children af 0000-3 er slettet
        self.assertEqual(len(result_df), 3)
        self.assertNotIn('Phantom', result_df['BOM Structure'].values)
        
if __name__ == '__main__':
    unittest.main() 