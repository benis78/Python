"""Excel håndtering i separate tråde"""

import os
import shutil
import threading
import queue
import tempfile
import sqlite3
import pandas as pd
import win32com.client
from pathlib import Path
from typing import Optional, Tuple, Dict, List
import config
from DrawingDatabase import DrawingDatabase

class ProcessingResult:
    """Resultat af databehandling"""
    def __init__(self, success: bool, message: str, data: Optional[dict] = None):
        self.success = success
        self.message = message
        self.data = data or {}

class DataProcessor(threading.Thread):
    """Tråd til pandas databehandling"""
    def __init__(self, input_file: Path, output_queue: queue.Queue, previous_file: Optional[Path] = None):
        super().__init__()
        self.input_file = input_file
        self.output_queue = output_queue
        self.previous_file = previous_file
        self.drawing_db = DrawingDatabase()
        self.temp_dir = None
        self.excel = None
        self.workbook = None
        
    def _create_temp_dir(self) -> Path:
        """Opret midlertidig mappe til billeder"""
        self.temp_dir = Path(tempfile.mkdtemp())
        return self.temp_dir
        
    def _cleanup_temp_dir(self):
        """Ryd op i midlertidig mappe"""
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
            
    def _extract_revision(self, filename: str) -> Tuple[str, str]:
        """Udtræk Part Number og REV fra filnavn"""
        if not filename or not isinstance(filename, str):
            return "None", ""
            
        # Split på " -- " hvis det findes
        parts = filename.split(" -- ")
        part_number = parts[0]
        
        # Find sidste bogstav som revision
        rev = ""
        if part_number and part_number[-1].isalpha():
            rev = part_number[-1]
            part_number = part_number[:-1]
            
        return part_number, rev
        
    def _extract_revision_from_partnumber(self, part_number: str) -> Tuple[str, str]:
        """Udtræk revision fra part number"""
        if not part_number or not isinstance(part_number, str):
            return "None", ""
            
        # Find sidste bogstav som revision
        rev = ""
        if part_number and part_number[-1].isalpha():
            rev = part_number[-1]
            part_number = part_number[:-1]
            
        return part_number, rev
        
    def _is_supplier_part(self, part_number: str) -> bool:
        """Check om part number er en leverandør del"""
        if not part_number or not isinstance(part_number, str):
            return False
            
        return any(part_number.startswith(prefix) for prefix in config.SUPPLIER_PREFIXES)
        
    def _process_bom_structure_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        """Behandl BOM Structure regler"""
        df = df.copy()
        
        # Fjern phantom rækker
        df = df[df['BOM Structure'] != 'Phantom']
        
        # Find og fjern børn af 0000-3 rækker og Inseparable
        parent_levels = df[
            (df['Part Number'].str.startswith('0000-3', na=False)) |
            (df['BOM Structure'] == 'Inseparable')
        ]['BOM Structure']
        
        child_mask = df['BOM Structure'].apply(
            lambda x: not any(str(x).startswith(str(p) + '.') for p in parent_levels)
        )
        df = df[child_mask]
        
        return df
        
    def _calculate_total_qty(self, df: pd.DataFrame) -> pd.DataFrame:
        """Beregn Total QTY baseret på BOM Structure"""
        df = df.copy()
        df['Total QTY'] = df['QTY']
        
        for idx, row in df.iterrows():
            if '.' in str(row['BOM Structure']):
                parent_level = '.'.join(str(row['BOM Structure']).split('.')[:-1])
                parent_qty = df[df['BOM Structure'] == parent_level]['QTY'].iloc[0]
                df.at[idx, 'Total QTY'] = row['QTY'] * parent_qty
                
        return df
        
    def _add_categories(self, df: pd.DataFrame) -> pd.DataFrame:
        """Tilføj kategorier baseret på part number"""
        df = df.copy()
        
        def get_category(part_number):
            if not part_number or not isinstance(part_number, str):
                return 'Unknown'
            for prefix, category in config.CATEGORIES.items():
                if part_number.startswith(prefix):
                    return category
            return 'Unknown'
            
        df['Category'] = df['Part Number'].apply(get_category)
        return df
        
    def _handle_thumbnails(self, df: pd.DataFrame) -> Dict[str, str]:
        """Håndter thumbnails fra Excel"""
        image_mapping = {}
        
        # Opret temp mappe hvis den ikke findes
        if not self.temp_dir:
            self._create_temp_dir()
            
        # Gem billeder fra Thumbnail kolonne
        for idx, row in df.iterrows():
            if pd.notna(row['Thumbnail']) and row['Thumbnail'] != '(NULL)':
                # Gem billede til temp mappe
                img_path = self.temp_dir / f"{row['Part Number']}.png"
                # TODO: Implementer gem billede logik
                image_mapping[row['Part Number']] = str(img_path)
                
        return image_mapping
        
    def _copy_drawings(self, df: pd.DataFrame, output_dir: Path) -> pd.DataFrame:
        """Kopier tegninger til kategori mapper"""
        df = df.copy()
        
        # Opret Drawing status kolonne
        df['Drawing Status'] = ''
        
        for idx, row in df.iterrows():
            drawings = self.drawing_db.find_drawings(row['Part Number'])
            
            if drawings:
                # Bestem kategori mappe
                category = row['Type'] if row['Category'] == 'Piping' else row['Category']
                category_dir = output_dir / category
                category_dir.mkdir(exist_ok=True)
                
                # Kopier tegninger
                dwg_found = False
                pdf_found = False
                dwg_rev = ''
                pdf_rev = ''
                
                for drawing in drawings:
                    if drawing['filename'].endswith('.dwg'):
                        dwg_found = True
                        dwg_rev = self._extract_revision_from_partnumber(
                            drawing['filename'][:-4]
                        )[1]
                        shutil.copy2(drawing['filepath'], category_dir)
                    elif drawing['filename'].endswith('.pdf'):
                        pdf_found = True
                        pdf_rev = self._extract_revision_from_partnumber(
                            drawing['filename'][:-4]
                        )[1]
                        shutil.copy2(drawing['filepath'], category_dir)
                
                # Sæt Drawing status
                if dwg_found and pdf_found:
                    if dwg_rev == pdf_rev:
                        df.at[idx, 'Drawing Status'] = config.DRAWING_STATUS['both_match']
                    else:
                        df.at[idx, 'Drawing Status'] = config.DRAWING_STATUS['both_mismatch']
                elif dwg_found:
                    df.at[idx, 'Drawing Status'] = config.DRAWING_STATUS['dwg_only']
                elif pdf_found:
                    df.at[idx, 'Drawing Status'] = config.DRAWING_STATUS['pdf_only']
                    
        return df
        
    def _create_partlist_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """Opret Partlist sheet"""
        # Kopier alle kolonner undtagen Item
        partlist_df = df.drop(columns=['Item']).copy()
        
        # Grupper efter Part Number og REV, sum Total QTY
        partlist_df = partlist_df.groupby(['Part Number', 'REV'], as_index=False).agg({
            'Total QTY': 'sum',
            **{col: 'first' for col in partlist_df.columns if col not in ['Part Number', 'REV', 'Total QTY']}
        })
        
        return partlist_df
        
    def _create_category_sheets(self, df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """Opret kategori sheets"""
        category_dfs = {}
        
        for category in df['Category'].unique():
            if category == 'Piping':
                # For Piping, brug Type som sheet navn
                for type_val in df[df['Category'] == 'Piping']['Type'].unique():
                    sheet_df = df[
                        (df['Category'] == 'Piping') &
                        (df['Type'] == type_val)
                    ].copy()
                    category_dfs[type_val] = sheet_df
            else:
                # For andre kategorier, brug Category som sheet navn
                sheet_df = df[df['Category'] == category].copy()
                category_dfs[category] = sheet_df
                
        return category_dfs
        
    def _create_compare_sheet(self, new_df: pd.DataFrame, old_df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
        """Opret Compare sheet"""
        compare_df = pd.DataFrame(columns=new_df.columns)
        formatting = {
            'new_rows': [],
            'changed_rows': [],
            'deleted_rows': []
        }
        
        # Find nye og ændrede rækker
        for idx, new_row in new_df.iterrows():
            old_row = old_df[
                (old_df['Part Number'] == new_row['Part Number']) &
                (old_df['REV'] == new_row['REV'])
            ]
            
            if len(old_row) == 0:
                # Ny række
                compare_df = compare_df.append(new_row)
                formatting['new_rows'].append(len(compare_df) - 1)
            else:
                old_row = old_row.iloc[0]
                # Check for ændringer
                changed = False
                for col in new_row.index:
                    if new_row[col] != old_row[col]:
                        changed = True
                        new_row[col] = f"{new_row[col]} ({old_row[col]})"
                
                if changed:
                    compare_df = compare_df.append(new_row)
                    formatting['changed_rows'].append(len(compare_df) - 1)
                    
        # Find slettede rækker
        for idx, old_row in old_df.iterrows():
            new_row = new_df[
                (new_df['Part Number'] == old_row['Part Number']) &
                (new_df['REV'] == old_row['REV'])
            ]
            
            if len(new_row) == 0:
                compare_df = compare_df.append(old_row)
                formatting['deleted_rows'].append(len(compare_df) - 1)
                
        return compare_df, formatting
        
    def _apply_excel_formatting(self, workbook) -> None:
        """Anvend Excel formatering"""
        try:
            for sheet in workbook.Sheets:
                # Indstil rækkehøjder
                sheet.Rows(1).RowHeight = config.EXCEL_FORMATTING['header_row_height']
                sheet.Rows(f"2:{sheet.UsedRange.Rows.Count}").RowHeight = (
                    config.EXCEL_FORMATTING['data_row_height']
                )
                
                # Indstil kolonnebredder
                for col, width in config.COLUMN_WIDTHS.items():
                    sheet.Columns(col).ColumnWidth = width
                    
                # Fed skrift i header
                if config.EXCEL_FORMATTING['header_row_bold']:
                    sheet.Rows(1).Font.Bold = True
                    
                # Tilføj filter
                if config.EXCEL_FORMATTING['auto_filter']:
                    sheet.Range(f"A1:{chr(64 + sheet.UsedRange.Columns.Count)}1").AutoFilter()
                    
                # Frys første række
                if config.EXCEL_FORMATTING['freeze_first_row']:
                    sheet.Rows(2).Select()
                    workbook.Windows(1).FreezePanes = True
                    
        except Exception as e:
            raise Exception(f"Fejl ved formatering: {str(e)}")
            
    def run(self):
        """Kør databehandling"""
        try:
            # Start Excel
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            
            # Load Excel fil
            self.workbook = self.excel.Workbooks.Open(str(self.input_file))
            df = pd.read_excel(self.input_file, sheet_name="BOM")
            
            # Opret output mappe
            output_dir = self.input_file.parent / f"{self.input_file.stem}_output"
            output_dir.mkdir(exist_ok=True)
            
            # Udtræk revision fra filnavn
            part_number, rev = self._extract_revision(self.input_file.stem)
            
            # Indsæt arrangement række
            arrangement_row = pd.DataFrame([{
                'Item': 0,
                'Part Number': part_number,
                'REV': rev,
                'BOM Structure': None,
                'Description': 'Arrangement Drawing',
                'QTY': 1,
                'D': 1,
                't': 1,
                'L': 1
            }])
            df = pd.concat([arrangement_row, df], ignore_index=True)
            
            # Behandl data
            df = self._process_bom_structure_rules(df)
            df = self._calculate_total_qty(df)
            df = self._add_categories(df)
            
            # Håndter thumbnails
            image_mapping = self._handle_thumbnails(df)
            
            # Kopier tegninger og tilføj Drawing status
            df = self._copy_drawings(df, output_dir)
            
            # Opret Partlist
            partlist_df = self._create_partlist_sheet(df)
            
            # Opret kategori sheets
            category_dfs = self._create_category_sheets(df)
            
            # Opret Compare sheet hvis previous_file findes
            compare_df = None
            compare_formatting = None
            if self.previous_file:
                old_df = pd.read_excel(self.previous_file, sheet_name="Partlist")
                compare_df, compare_formatting = self._create_compare_sheet(partlist_df, old_df)
            
            # Gem resultat
            output_file = output_dir / self.input_file.name
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Gem sheets i korrekt rækkefølge
                df.to_excel(writer, sheet_name='BOM', index=False)
                partlist_df.to_excel(writer, sheet_name='Partlist', index=False)
                
                if compare_df is not None:
                    old_rev = self._extract_revision(self.previous_file.stem)[1]
                    new_rev = rev
                    compare_df.to_excel(
                        writer,
                        sheet_name=f'Compare OLD{old_rev}-NEW{new_rev}',
                        index=False
                    )
                    
                # Gem kategori sheets i alfabetisk rækkefølge
                for category in sorted(category_dfs.keys()):
                    category_dfs[category].to_excel(
                        writer,
                        sheet_name=category,
                        index=False
                    )
                    
            # Åbn den gemte fil for formatering
            workbook = self.excel.Workbooks.Open(str(output_file))
            
            # Anvend formatering
            self._apply_excel_formatting(workbook)
            
            # Gem og luk
            workbook.Save()
            workbook.Close()
            
            # Send resultat
            self.output_queue.put(ProcessingResult(
                success=True,
                message="Data processing completed successfully",
                data={
                    'output_folder': str(output_dir),
                    'df': df,
                    'partlist_df': partlist_df,
                    'category_dfs': category_dfs,
                    'compare_df': compare_df,
                    'compare_formatting': compare_formatting
                }
            ))
            
        except Exception as e:
            self.output_queue.put(ProcessingResult(
                success=False,
                message=f"Error during data processing: {str(e)}"
            ))
            
        finally:
            # Cleanup
            if self.workbook:
                try:
                    self.workbook.Close(SaveChanges=False)
                except:
                    pass
            if self.excel:
                try:
                    self.excel.Quit()
                except:
                    pass
            self._cleanup_temp_dir() 