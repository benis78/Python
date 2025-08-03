"""
Test script for image handling
"""
import logging
from pathlib import Path
import win32com.client
import pythoncom
from excel_handler import ExcelHandler
import os

def inspect_excel_file(file_path: Path):
    """Undersøg Excel filens struktur"""
    logging.info(f"Inspicerer Excel fil: {file_path}")
    
    # Start Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    
    try:
        # Åbn workbook
        wb = excel.Workbooks.Open(str(file_path.absolute()))
        sheet = wb.Sheets(1)
        
        # Find kolonner
        for col in range(1, sheet.UsedRange.Columns.Count + 1):
            header = sheet.Cells(1, col).Value
            logging.info(f"Kolonne {col}: {header}")
            
        # Tjek for shapes/billeder
        logging.info(f"Antal shapes: {sheet.Shapes.Count}")
        for shape in sheet.Shapes:
            logging.info(f"Shape: Type={shape.Type}, Name={shape.Name}")
            
    finally:
        # Luk Excel
        excel.Quit()

def test_image_handling():
    # Setup logging
    logging.basicConfig(level=logging.INFO)
    
    # Input og output stier
    current_dir = Path(os.getcwd())
    input_file = current_dir / "4003-02.1-A01-- - BOM.xlsx"
    output_file = input_file.parent / f"{input_file.stem}_ImageTest.xlsx"
    
    if not input_file.exists():
        logging.error(f"Kunne ikke finde input fil: {input_file}")
        return
    
    logging.info(f"Testing image handling med fil: {input_file}")
    
    try:
        # Undersøg Excel fil først
        inspect_excel_file(input_file)
        
        # Initialiser Excel handler
        handler = ExcelHandler(str(input_file))
        
        # Load workbook
        logging.info("Indlæser workbook...")
        handler.load_workbook()
        
        # Udtræk billeder
        logging.info("Udtrækker billeder...")
        handler._extract_images()
        
        # Log antal fundne billeder
        logging.info(f"Fandt {len(handler.image_mapping)} billeder")
        for part_num, img_path in handler.image_mapping.items():
            logging.info(f"Billede for {part_num}: {img_path}")
            
        # Gem workbook med billeder
        logging.info(f"Gemmer workbook til: {output_file}")
        handler.save_workbook(str(output_file))
        
        logging.info("Test færdig!")
        
    except Exception as e:
        logging.exception("Fejl under test:")
        raise

if __name__ == "__main__":
    test_image_handling() 