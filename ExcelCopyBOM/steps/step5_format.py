"""
TRIN 5: Formatering af alle ark
Håndterer formatering af alle ark i workbook
"""

import logging

class WorksheetFormatter:
    def __init__(self, workbook):
        self.logger = logging.getLogger('ExcelCopyBOM.Formatter')
        self.workbook = workbook
        
    def _format_worksheet(self, sheet):
        """Formaterer et enkelt worksheet"""
        try:
            # Indstil kolonne bredder (pixel / 7)
            column_widths = {
                "Item": 10,  # 70 pixels
                "Part Number": 20,  # 140 pixels
                "Description": 60,  # 420 pixels
                "Quantity": 10,  # 70 pixels
                "Unit": 8,  # 56 pixels
                "Weight": 10,  # 70 pixels
                "Material": 15,  # 105 pixels
                "Surface Treatment": 20,  # 140 pixels
                "REV": 8,  # 56 pixels
                "BOM Structure": 15,  # 105 pixels
                "Drawing": 12  # 84 pixels
            }
            
            # Find kolonner
            for i in range(1, sheet.UsedRange.Columns.Count + 1):
                header = str(sheet.Cells(1, i).Value).strip()
                if header in column_widths:
                    sheet.Columns(i).ColumnWidth = column_widths[header]
                    
            # Indstil række højder
            sheet.Rows(1).RowHeight = 20  # Header række
            if sheet.UsedRange.Rows.Count > 1:
                sheet.Range(f"2:{sheet.UsedRange.Rows.Count}").RowHeight = 15
                
            # Frys første række
            sheet.Application.ActiveWindow.SplitRow = 1
            sheet.Application.ActiveWindow.FreezePanes = True
            
            # Tilføj filter
            sheet.UsedRange.AutoFilter()
            
            # Juster tekst
            sheet.UsedRange.HorizontalAlignment = -4131  # xlLeft
            sheet.UsedRange.VerticalAlignment = -4108  # xlCenter
            
            # Formatér header
            header_range = sheet.Range("1:1")
            header_range.Font.Bold = True
            header_range.Interior.Color = 15921906  # RGB(242, 242, 242)
            
            # Tilføj borders
            sheet.UsedRange.Borders.LineStyle = 1  # xlContinuous
            sheet.UsedRange.Borders.Weight = 2  # xlThin
            
            # Wrap text i Description
            for i in range(1, sheet.UsedRange.Columns.Count + 1):
                if str(sheet.Cells(1, i).Value).strip() == "Description":
                    sheet.Columns(i).WrapText = True
                    break
                    
        except Exception as e:
            self.logger.error(f"Fejl under formatering af {sheet.Name}: {str(e)}")
            
    def format_all_sheets(self) -> bool:
        """
        Udfører TRIN 5: Formatering af alle ark
        """
        try:
            self.logger.info("Starter formatering af alle ark")
            
            # Formatér alle ark
            for sheet in self.workbook.Sheets:
                self.logger.debug(f"Formaterer ark: {sheet.Name}")
                self._format_worksheet(sheet)
                
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under formatering: {str(e)}")
            return False 