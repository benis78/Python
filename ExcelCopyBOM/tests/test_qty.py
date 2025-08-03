"""
Test af QTY beregning med rigtig BOM fil
"""
import unittest
import pandas as pd
from pathlib import Path
import logging
from ExcelCopyBOM.threaded_excel_handler import DataProcessor
import queue

class TestQtyCalculation(unittest.TestCase):
    def setUp(self):
        self.bom_file = Path("C:/Coding/Python/ExcelCopyBOM/4003-02.1-A01-- - BOM.xlsx")
        self.output_queue = queue.Queue()
        self.processor = DataProcessor(self.bom_file, self.output_queue)
        
    def test_qty_calculation_with_real_bom(self):
        """Test QTY beregning med rigtig BOM fil"""
        # Load Excel fil
        df = pd.read_excel(self.bom_file, engine='openpyxl')
        
        # Print original QTY værdier
        print("\nOriginale QTY værdier:")
        for idx, row in df.iterrows():
            if pd.notna(row['BOM Structure']):
                print(f"Row {idx+1}: Structure={row['BOM Structure']}, QTY={row['QTY']}")
        
        # Beregn Total QTY
        df = self.processor._calculate_total_qty(df)
        
        # Print resultater
        print("\nBeregnede Total QTY værdier:")
        for idx, row in df.iterrows():
            if pd.notna(row['BOM Structure']):
                print(f"Row {idx+1}: Structure={row['BOM Structure']}, QTY={row['QTY']}, Total QTY={row['Total QTY']}")
        
        # Verificer at Total QTY er beregnet korrekt for nogle kendte værdier
        # Eksempel: Hvis vi har en parent med QTY=2 og et child med QTY=3,
        # så skal child's Total QTY være 6
        
        # Find et eksempel med parent/child relation
        for idx, row in df.iterrows():
            if pd.notna(row['BOM Structure']) and '.' in str(row['BOM Structure']):
                parent_structure = '.'.join(str(row['BOM Structure']).split('.')[:-1])
                parent_row = df[df['BOM Structure'] == parent_structure].iloc[0]
                
                child_qty = float(row['QTY'])
                parent_qty = float(parent_row['QTY'])
                total_qty = float(row['Total QTY'])
                
                print(f"\nTest af parent/child relation:")
                print(f"Parent: Structure={parent_structure}, QTY={parent_qty}")
                print(f"Child: Structure={row['BOM Structure']}, QTY={child_qty}, Total QTY={total_qty}")
                
                expected_total = child_qty * parent_qty
                self.assertAlmostEqual(total_qty, expected_total, places=2,
                    msg=f"Forventet Total QTY {expected_total} men fik {total_qty}")
                break

if __name__ == '__main__':
    unittest.main() 