"""
Image handling for Excel using win32com
"""
import win32com.client
import pythoncom
import threading
import logging
from pathlib import Path
import time
from typing import Dict, Optional

class ExcelImageHandler:
    def __init__(self):
        self._excel = None
        self._workbook = None
        self._worksheet = None

    def _init_excel_com(self):
        """Initialiserer Excel COM objektet"""
        logging.info("Initialiserer Excel COM objekt")
        pythoncom.CoInitialize()
        self._excel = win32com.client.Dispatch("Excel.Application")
        self._excel.Visible = False
        self._excel.DisplayAlerts = False

    def _cleanup_excel_com(self):
        """Lukker Excel ned og frigiver COM objektet"""
        if self._excel:
            logging.info("Lukker Excel ned")
            self._excel.Quit()
            pythoncom.CoUninitialize()

    def process_images(self, output_file: str, image_mapping: Dict[str, str]) -> None:
        """
        Indsæt billeder i Excel fil.
        Args:
            output_file: Sti til output Excel fil
            image_mapping: Dictionary med part number -> billede sti
        """
        try:
            # Initialiser Excel
            self._init_excel_com()
            
            # Åbn workbook
            logging.info(f"Åbner workbook for billede indsættelse: {output_file}")
            workbook = self._excel.Workbooks.Open(output_file)
            sheet = workbook.Sheets(1)
            
            # Find Thumbnail og Part Number kolonner
            thumbnail_col = None
            part_number_col = None
            for col in range(1, sheet.UsedRange.Columns.Count + 1):
                if sheet.Cells(1, col).Value == "Thumbnail":
                    thumbnail_col = col
                elif sheet.Cells(1, col).Value == "Part Number":
                    part_number_col = col
                    
            if not thumbnail_col or not part_number_col:
                raise ValueError("Kunne ikke finde Thumbnail eller Part Number kolonne")
                
            # Gennemgå hver række
            for row in range(2, sheet.UsedRange.Rows.Count + 1):
                part_number = str(sheet.Cells(row, part_number_col).Value)
                if part_number in image_mapping:
                    # Slet eksisterende billeder i cellen
                    cell = sheet.Cells(row, thumbnail_col)
                    for shape in sheet.Shapes:
                        if (shape.Left >= cell.Left and 
                            shape.Left <= cell.Left + cell.Width and
                            shape.Top >= cell.Top and 
                            shape.Top <= cell.Top + cell.Height):
                            shape.Delete()
                    
                    # Indsæt nyt billede
                    image_path = image_mapping[part_number]
                    if Path(image_path).exists():
                        # Indsæt billede
                        picture = sheet.Shapes.AddPicture(
                            image_path,
                            LinkToFile=False,
                            SaveWithDocument=True,
                            Left=cell.Left,
                            Top=cell.Top,
                            Width=0,  # 0 = auto width
                            Height=0   # 0 = auto height
                        )
                        
                        # Sæt billede højde til 2.38 cm (konverter fra cm til points)
                        target_height = 2.38 * 28.3465  # 1 cm = 28.3465 points
                        scale_factor = target_height / picture.Height
                        
                        picture.Height = target_height
                        picture.Width = picture.Width * scale_factor
                        
                        # Centrer i cellen
                        picture.Left = cell.Left + (cell.Width - picture.Width) / 2
                        picture.Top = cell.Top + (cell.Height - picture.Height) / 2
                        
                        logging.info(f"Indsat billede for {part_number} i række {row}")
                        
            # Gem ændringer
            workbook.Save()
            
        except Exception as e:
            logging.error(f"Fejl ved indsættelse af billeder: {str(e)}")
            raise
            
        finally:
            self._cleanup_excel_com()