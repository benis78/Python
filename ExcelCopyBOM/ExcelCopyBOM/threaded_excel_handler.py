"""
Threaded Excel handler med parallel processering af data, billeder og filer
"""
import threading
import queue
import logging
import pandas as pd
import tempfile
import win32com.client
import pythoncom
import win32clipboard
from PIL import ImageGrab
from pathlib import Path
from typing import Dict, Optional, List
from dataclasses import dataclass
# Ændret fra relative til absolute imports
from image_handler import ExcelImageHandler
from Categories import PartNumberParser
from DrawingDatabase import DrawingDatabase

@dataclass
class ProcessingResult:
    """Data klasse til at holde resultater fra forskellige tråde"""
    success: bool
    error: Optional[str] = None
    data: Optional[Dict] = None

class DataProcessor(threading.Thread):
    """Tråd til pandas databehandling"""
    def __init__(self, input_file: Path, output_queue: queue.Queue, previous_file: Optional[Path] = None):
        super().__init__()
        self.input_file = input_file
        self.previous_file = previous_file
        self.output_queue = output_queue
        self.result = ProcessingResult(success=False)
        self.part_number_parser = PartNumberParser()
        self.db = DrawingDatabase()
        self.compare_formatting = None  # Gemmer formatering info til Compare sheet
        
    def _extract_revision(self, filename: str) -> tuple[str, str]:
        """
        Udtræk Part Number og REV fra filnavn
        Format: XXXX-XX.X-XXX -- Description.xlsx
        REV er sidste bogstav før " - "
        """
        # Find position af " - "
        dash_pos = filename.find(" - ")
        if dash_pos == -1:
            return filename, ""
            
        # Part number er alt før " - "
        part_number = filename[:dash_pos]
        
        # REV er sidste bogstav i part number
        rev = ""
        for char in reversed(part_number):
            if char.isalpha():
                rev = char
                # Fjern REV fra part number
                part_number = part_number.replace(rev, "")
                break
                
        return part_number.strip(), rev.strip()
        
    def _extract_revision_from_partnumber(self, part_number: str) -> tuple[str, str]:
        """
        Udtræk revision fra part number og returner rent part number
        F.eks.: "1234-56-A01A" -> ("1234-56-A01", "A")
        """
        if not isinstance(part_number, str):
            return str(part_number), ""
            
        rev = ""
        clean_part_number = part_number
        
        # Find sidste bogstav i part number
        for char in reversed(part_number):
            if char.isalpha():
                rev = char
                # Fjern REV fra part number
                clean_part_number = part_number.replace(rev, "")
                break
                
        return clean_part_number.strip(), rev.strip()
        
    def _is_supplier_part(self, part_number: str) -> bool:
        """Check om part number er en supplier part"""
        if not isinstance(part_number, str):
            return False
            
        supplier_prefixes = ['0000-700', '0000-701', '0000-702']
        return any(part_number.startswith(prefix) for prefix in supplier_prefixes)
        
    def _calculate_total_qty(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Beregn Total QTY ved at gange QTY med parent QTY.
        Håndterer:
        - Tekst-formaterede QTY værdier
        - Manglende BOM Structure
        - Manglende parent rows
        - NaN/None værdier i QTY
        """
        # Konverter QTY kolonne til numerisk, håndter tekst-formatering
        df['QTY'] = pd.to_numeric(df['QTY'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        
        # Tilføj Total QTY kolonne hvis den ikke findes
        if 'Total QTY' not in df.columns:
            df['Total QTY'] = pd.NA

        # Initialiser Total QTY med QTY værdier
        df['Total QTY'] = df['QTY']
        
        try:
            # Opbyg structure map for hurtigere parent lookup
            structure_map = {}
            for idx, row in df.iterrows():
                if pd.notna(row['BOM Structure']):
                    structure_map[str(row['BOM Structure'])] = {
                        'idx': idx,
                        'qty': float(row['QTY'])  # Nu er QTY allerede konverteret til float
                    }
            
            # Gennemgå hver række og beregn Total QTY
            for idx, row in df.iterrows():
                try:
                    if pd.notna(row['BOM Structure']):
                        structure = str(row['BOM Structure'])
                        level = len(structure.split('.'))
                        
                        if level > 1:
                            # Find parent structure
                            parent_structure = '.'.join(structure.split('.')[:-1])
                            
                            # Find parent i structure map
                            if parent_structure in structure_map:
                                parent_idx = structure_map[parent_structure]['idx']
                                parent_total_qty = df.at[parent_idx, 'Total QTY']
                                
                                # Beregn Total QTY (QTY er nu garanteret numerisk)
                                df.at[idx, 'Total QTY'] = row['QTY'] * parent_total_qty
                                
                                # Opdater structure map
                                structure_map[structure]['qty'] = df.at[idx, 'Total QTY']
                            else:
                                logging.warning(f"Parent structure {parent_structure} ikke fundet for række {idx}")
                                
                except Exception as e:
                    logging.warning(f"Fejl ved beregning af Total QTY for række {idx}: {str(e)}")
                    continue
                    
        except Exception as e:
            logging.error(f"Fejl i Total QTY beregning: {str(e)}")
            
        # Afrund Total QTY til 2 decimaler
        df['Total QTY'] = df['Total QTY'].round(2)
            
        return df
        
    def _add_categories(self, df: pd.DataFrame) -> pd.DataFrame:
        """Tilføj kategori-kolonne baseret på part numbers"""
        df['Category'] = df['Part Number'].apply(
            lambda x: self.part_number_parser.find_category(str(x)) if pd.notna(x) else ''
        )
        return df
        
    def _process_bom_structure_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Håndter BOM Structure regler:
        - Slet children hvis parent er "Inseparable" eller starter med "0000-3"
        - Slet rækker med "Phantom"
        """
        rows_to_delete = set()

        # Opbyg parent/child relationer
        structure_map = {}
        for idx, row in df.iterrows():
            if pd.notna(row['BOM Structure']):
                structure = str(row['BOM Structure'])
                structure_map[structure] = {
                    'idx': idx,
                    'part_number': str(row['Part Number']),
                    'structure': structure
                }

        # Find rækker der skal slettes
        for idx, row in df.iterrows():
            if pd.notna(row['BOM Structure']):
                structure = str(row['BOM Structure'])
                
                # Slet hvis Phantom
                if structure.lower() == 'phantom':
                    rows_to_delete.add(idx)
                    logging.info(f"Markerer Phantom række {idx} til sletning")
                    continue

                # Find parent for denne række
                if '.' in structure:
                    parent_structure = '.'.join(structure.split('.')[:-1])
                    if parent_structure in structure_map:
                        parent = structure_map[parent_structure]
                        parent_row = df.iloc[parent['idx']]

                        # Slet hvis parent er Inseparable
                        if str(parent_row['BOM Structure']).lower() == 'inseparable':
                            rows_to_delete.add(idx)
                            logging.info(f"Markerer child {idx} af Inseparable parent til sletning")

                        # Slet hvis parent starter med 0000-3
                        elif str(parent_row['Part Number']).startswith('0000-3'):
                            rows_to_delete.add(idx)
                            logging.info(f"Markerer child {idx} af 0000-3 parent til sletning")

        # Slet markerede rækker
        if rows_to_delete:
            logging.info(f"Sletter {len(rows_to_delete)} rækker baseret på BOM Structure regler")
            df = df.drop(index=rows_to_delete).reset_index(drop=True)

        return df

    def _copy_drawings_to_categories(self, df: pd.DataFrame, output_dir: Path) -> None:
        """
        Kopier tegninger til kategori-mapper baseret på part number kategorier.
        Opretter følgende mappestruktur:
        output_dir/
            Category1/
                DWG/
                PDF/
            Category2/
                DWG/
                PDF/
            ...
        """
        try:
            # Opret output directory hvis det ikke findes
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Gennemgå hver række i DataFrame
            for idx, row in df.iterrows():
                if pd.isna(row['Part Number']):
                    continue
                    
                part_number = str(row['Part Number'])
                category = row.get('Category', 'Uncategorized')
                
                # Opret kategori mapper
                category_dir = output_dir / category
                dwg_dir = category_dir / 'DWG'
                pdf_dir = category_dir / 'PDF'
                
                for dir_path in [category_dir, dwg_dir, pdf_dir]:
                    dir_path.mkdir(parents=True, exist_ok=True)
                
                # Find tegninger i databasen
                drawings = self.db.find_drawings(part_number)
                if not drawings:
                    logging.warning(f"Ingen tegninger fundet for {part_number}")
                    continue
                    
                # Kopier hver tegning til den rigtige mappe
                for drawing in drawings:
                    source_path = Path(drawing['filepath'])
                    if not source_path.exists():
                        logging.warning(f"Kildefil findes ikke: {source_path}")
                        continue
                        
                    # Bestem målmappe baseret på filtype
                    if source_path.suffix.lower() == '.dwg':
                        target_dir = dwg_dir
                    elif source_path.suffix.lower() == '.pdf':
                        target_dir = pdf_dir
                    else:
                        logging.warning(f"Ukendt filtype for {source_path}")
                        continue
                    
                    # Kopier fil
                    target_path = target_dir / source_path.name
                    try:
                        import shutil
                        shutil.copy2(source_path, target_path)
                        logging.info(f"Kopieret {source_path.name} til {target_path}")
                    except Exception as e:
                        logging.error(f"Fejl ved kopiering af {source_path} til {target_path}: {str(e)}")
                        
        except Exception as e:
            logging.error(f"Fejl i _copy_drawings_to_categories: {str(e)}")
            raise

    def _update_revisions_from_drawings(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Opdater REV kolonne baseret på seneste tegninger i databasen.
        Kun opdater hvis tegningen har en nyere revision.
        """
        try:
            for idx, row in df.iterrows():
                if pd.isna(row['Part Number']):
                    continue
                    
                part_number = str(row['Part Number'])
                current_rev = str(row['REV']) if pd.notna(row['REV']) else ''
                
                # Find tegninger for dette part number
                drawings = self.db.find_drawings(part_number)
                if not drawings:
                    continue
                
                # Find seneste revision fra tegninger
                latest_rev = ''
                for drawing in drawings:
                    # Udtræk revision fra filnavn
                    filename = Path(drawing['filename']).stem
                    _, drawing_rev = self._extract_revision(filename)
                    
                    # Opdater latest_rev hvis denne tegning har en nyere revision
                    if drawing_rev and (not latest_rev or drawing_rev > latest_rev):
                        latest_rev = drawing_rev
                
                # Opdater REV hvis vi fandt en nyere revision
                if latest_rev and (not current_rev or latest_rev > current_rev):
                    df.at[idx, 'REV'] = latest_rev
                    logging.info(f"Opdateret revision for {part_number} fra {current_rev} til {latest_rev}")
                    
        except Exception as e:
            logging.error(f"Fejl ved opdatering af revisioner: {str(e)}")
            
        return df

    def _add_drawing_status(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Tilføj Drawing status kolonne med følgende værdier:
        - DWG~PDF: Både DWG og PDF findes
        - DWG!PDF: Kun DWG findes
        - PDF: Kun PDF findes
        - DWG: Kun DWG findes
        - tom: Ingen tegninger findes
        """
        try:
            # Tilføj Drawing status kolonne
            df['Drawing status'] = ''
            
            for idx, row in df.iterrows():
                if pd.isna(row['Part Number']):
                    continue
                    
                part_number = str(row['Part Number'])
                
                # Find tegninger i databasen
                drawings = self.db.find_drawings(part_number)
                if not drawings:
                    continue
                    
                # Check hvilke filtyper der findes
                has_dwg = any(Path(d['filepath']).suffix.lower() == '.dwg' for d in drawings)
                has_pdf = any(Path(d['filepath']).suffix.lower() == '.pdf' for d in drawings)
                
                # Bestem status
                if has_dwg and has_pdf:
                    status = 'DWG~PDF'
                elif has_dwg:
                    status = 'DWG!PDF'
                elif has_pdf:
                    status = 'PDF'
                else:
                    status = ''
                    
                df.at[idx, 'Drawing status'] = status
                
        except Exception as e:
            logging.error(f"Fejl ved tilføjelse af Drawing status: {str(e)}")
            
        return df

    def _create_partlist_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Opret Partlist sheet ved at:
        1. Kopiere header row (undtagen Item)
        2. Aggregere identiske Part Number/REV rækker
        3. Summere Total QTY for identiske items
        4. Kopiere alle unikke Part Number/REV kombinationer
        """
        try:
            # Kopier relevante kolonner (alle undtagen Item)
            columns_to_copy = [col for col in df.columns if col.lower() != 'item']
            partlist_df = df[columns_to_copy].copy()
            
            # Fjern tomme rækker og rækker uden part number
            partlist_df = partlist_df.dropna(subset=['Part Number'])
            
            # Grupper efter Part Number og REV, sumér Total QTY
            group_columns = ['Part Number', 'REV']
            agg_dict = {
                'Total QTY': 'sum',
                'Description': 'first',
                'Material': 'first',
                'Title': 'first',
                'D': 'first',
                't': 'first',
                'L': 'first',
                'Keywords': 'first',
                'Type': 'first',
                'Category': 'first',
                'Drawing status': 'first'
            }
            
            # Tilføj alle andre kolonner som 'first'
            for col in partlist_df.columns:
                if col not in agg_dict and col not in group_columns:
                    agg_dict[col] = 'first'
            
            # Aggreger data
            partlist_df = partlist_df.groupby(group_columns, as_index=False).agg(agg_dict)
            
            # Sorter efter Part Number og REV
            partlist_df = partlist_df.sort_values(by=['Part Number', 'REV'])
            
            # Reset index
            partlist_df = partlist_df.reset_index(drop=True)
            
            logging.info(f"Oprettet Partlist med {len(partlist_df)} unikke dele")
            
            return partlist_df
            
        except Exception as e:
            logging.error(f"Fejl ved oprettelse af Partlist: {str(e)}")
            raise

    def _create_category_sheets(self, df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """
        Opret sheets baseret på kategorier.
        Speciel håndtering af Piping kategori (brug Type værdi i stedet).
        Returns:
            Dict med kategori -> DataFrame
        """
        try:
            category_sheets = {}
            
            # Kopier header row (undtagen Item)
            columns_to_copy = [col for col in df.columns if col.lower() != 'item']
            
            # Gruppér efter kategori
            for category in df['Category'].unique():
                if pd.isna(category):
                    continue
                    
                # Filtrer rows for denne kategori
                if category.lower() == 'piping':
                    # For Piping, gruppér efter Type
                    piping_df = df[df['Category'] == category].copy()
                    for pipe_type in piping_df['Type'].unique():
                        if pd.notna(pipe_type):
                            type_df = piping_df[piping_df['Type'] == pipe_type].copy()
                            # Inkluder children rows baseret på BOM Structure
                            for idx, row in type_df.iterrows():
                                if pd.notna(row['BOM Structure']):
                                    structure = str(row['BOM Structure'])
                                    # Find alle children
                                    children = df[df['BOM Structure'].apply(
                                        lambda x: pd.notna(x) and str(x).startswith(structure + '.')
                                    )]
                                    type_df = pd.concat([type_df, children]).drop_duplicates()
                            
                            sheet_name = f"Piping - {pipe_type}"
                            category_sheets[sheet_name] = type_df[columns_to_copy]
                else:
                    # For andre kategorier, tag alle rows med denne kategori
                    category_df = df[df['Category'] == category].copy()
                    category_sheets[category] = category_df[columns_to_copy]
                    
            # Sorter DataFrames
            for category, cat_df in category_sheets.items():
                category_sheets[category] = cat_df.sort_values(by=['Part Number', 'REV']).reset_index(drop=True)
                
            logging.info(f"Oprettet {len(category_sheets)} kategori sheets")
            
            return category_sheets
            
        except Exception as e:
            logging.error(f"Fejl ved oprettelse af kategori sheets: {str(e)}")
            raise

    def _create_compare_sheet(self, current_df: pd.DataFrame, previous_df: pd.DataFrame) -> pd.DataFrame:
        """
        Opret Compare sheet ved at sammenligne current og previous Partlist.
        Format: "Compare OLDREV-NEWREV"
        Regler:
        - Kopier NEW rows med højere REV
        - Kopier OLD rows der mangler i NEW
        - Vis ændrede værdier med gamle værdier i parentes
        - Farvekodning:
          - Lysegrøn + fed for nye rækker
          - Lysgul + fed for ændrede værdier
          - Lysrød for slettede rækker
        """
        try:
            # Kopier header row (undtagen Item)
            columns_to_copy = [col for col in current_df.columns if col.lower() != 'item']
            
            # Opret tom DataFrame til Compare sheet
            compare_df = pd.DataFrame(columns=columns_to_copy)
            
            # Find OLD og NEW revision fra filnavne
            _, old_rev = self._extract_revision(self.previous_file.stem)
            _, new_rev = self._extract_revision(self.input_file.stem)
            
            # Tilføj Style kolonne til at holde styr på formatering
            compare_df['_Style'] = ''
            
            # Sammenlign hver række
            for idx, new_row in current_df.iterrows():
                part_number = new_row['Part Number']
                new_rev = new_row['REV']
                
                # Find matchende række i previous_df
                old_rows = previous_df[previous_df['Part Number'] == part_number]
                
                if len(old_rows) == 0:
                    # Ny del - markér med lysegrøn
                    new_row['_Style'] = 'new'
                    compare_df = pd.concat([compare_df, pd.DataFrame([new_row])], ignore_index=True)
                else:
                    old_row = old_rows.iloc[0]
                    old_rev = old_row['REV']
                    
                    if new_rev > old_rev:
                        # Opdateret revision - markér ændrede celler med lysgul
                        new_row['_Style'] = 'changed'
                        # Tilføj gamle værdier i parentes for ændrede felter
                        for col in columns_to_copy:
                            if new_row[col] != old_row[col]:
                                new_row[col] = f"{new_row[col]} ({old_row[col]})"
                        compare_df = pd.concat([compare_df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Find slettede rækker (findes i OLD men ikke i NEW)
            for idx, old_row in previous_df.iterrows():
                part_number = old_row['Part Number']
                if len(current_df[current_df['Part Number'] == part_number]) == 0:
                    # Slettet del - markér med lysrød
                    old_row['_Style'] = 'deleted'
                    compare_df = pd.concat([compare_df, pd.DataFrame([old_row])], ignore_index=True)
            
            # Sorter efter Part Number
            compare_df = compare_df.sort_values(by=['Part Number']).reset_index(drop=True)
            
            # Gem formatering information til senere brug
            self.compare_formatting = {
                'new_rows': compare_df[compare_df['_Style'] == 'new'].index.tolist(),
                'changed_rows': compare_df[compare_df['_Style'] == 'changed'].index.tolist(),
                'deleted_rows': compare_df[compare_df['_Style'] == 'deleted'].index.tolist()
            }
            
            # Fjern _Style kolonne før returnering
            compare_df = compare_df.drop('_Style', axis=1)
            
            logging.info(f"Oprettet Compare sheet (OLDREV={old_rev}, NEWREV={new_rev})")
            return compare_df
            
        except Exception as e:
            logging.error(f"Fejl ved oprettelse af Compare sheet: {str(e)}")
            raise

    def _apply_formatting(self, workbook) -> None:
        """
        Anvend formatering på alle ark:
        - Række højder
        - Kolonne bredder
        - Fed skrift i header
        - Filtre
        - Frys øverste række
        """
        try:
            # Konverter pixel til points (1 pixel ≈ 0.75 points)
            header_height = 26 * 0.75  # 26px -> points
            data_row_height = 91 * 0.75  # 91px -> points
            
            # Kolonne bredder i points (1 pixel ≈ 0.75 points)
            column_widths = {
                'A': 52 * 0.75,   # Item
                'B': 152 * 0.75,  # Part Number
                'C': 47 * 0.75,   # Rev
                'D': 111 * 0.75,  # Description 1
                'E': 93 * 0.75,   # Thumbnail
                'F': 115 * 0.75,  # BOM Structure
                'G': 423 * 0.75,  # Description
                'H': 135 * 0.75,  # Material
                'I': 173 * 0.75,  # Title
                'J': 48 * 0.75,   # QTY
                'K': 82 * 0.75,   # Total QTY
                'L': 39 * 0.75,   # Diameter
                'M': 39 * 0.75,   # Thickness
                'N': 39 * 0.75,   # Length
                'O': 200 * 0.75,  # Keywords
                'P': 180 * 0.75,  # Type
                'Q': 125 * 0.75,  # Category
                'R': 94 * 0.75    # Drawings
            }
            
            # Anvend formatering på hvert ark
            for sheet in workbook.Sheets:
                # Indstil række højder
                sheet.Rows(1).RowHeight = header_height  # Header række
                sheet.Rows(f"2:{sheet.UsedRange.Rows.Count}").RowHeight = data_row_height  # Data rækker
                
                # Indstil kolonne bredder
                for col, width in column_widths.items():
                    sheet.Columns(col).ColumnWidth = width
                
                # Fed skrift i header
                sheet.Rows(1).Font.Bold = True
                
                # Tilføj filtre til header række
                sheet.Range(f"A1:{chr(64 + sheet.UsedRange.Columns.Count)}1").AutoFilter()
                
                # Frys øverste række
                sheet.Rows(2).Select()
                workbook.Windows(1).FreezePanes = True
                
            logging.info("Formatering anvendt på alle ark")
            
        except Exception as e:
            logging.error(f"Fejl ved anvendelse af formatering: {str(e)}")
            raise

    def run(self):
        try:
            logging.info("Starter data processing")
            
            # Load Excel fil med pandas
            logging.info(f"Indlæser Excel fil: {self.input_file}")
            df = pd.read_excel(self.input_file, engine='openpyxl')
            
            # Gem original columns
            original_columns = df.columns.tolist()
            
            # Extract Part Number og REV fra filnavn
            part_number, rev = self._extract_revision(self.input_file.stem)
            logging.info(f"Extracted from filename: Part Number={part_number}, REV={rev}")
            
            # Indsæt arrangement row i row 2
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
            
            df = pd.concat([df.iloc[:1], arrangement_row, df.iloc[1:]], ignore_index=True)
            
            # Slet supplier parts
            initial_rows = len(df)
            df = df[~df['Part Number'].apply(self._is_supplier_part)]
            deleted_rows = initial_rows - len(df)
            logging.info(f"Deleted {deleted_rows} supplier part rows")
            
            # Flyt revision bogstaver til REV kolonne
            for idx, row in df.iterrows():
                if pd.notna(row['Part Number']):
                    clean_part_number, part_rev = self._extract_revision_from_partnumber(row['Part Number'])
                    df.at[idx, 'Part Number'] = clean_part_number
                    if part_rev and (pd.isna(row['REV']) or not row['REV']):
                        df.at[idx, 'REV'] = part_rev

            # Håndter BOM Structure regler
            df = self._process_bom_structure_rules(df)
                        
            # Beregn Total QTY
            df = self._calculate_total_qty(df)
            
            # Tilføj kategorier
            df = self._add_categories(df)
            logging.info("Added categories based on part numbers")
            
            # Kopier tegninger til kategori-mapper
            output_dir = self.input_file.parent / f"{self.input_file.stem}_drawings"
            self._copy_drawings_to_categories(df, output_dir)
            logging.info(f"Kopieret tegninger til {output_dir}")
            
            # Opdater revisioner fra tegninger
            df = self._update_revisions_from_drawings(df)
            logging.info("Opdateret revisioner fra tegninger")
            
            # Tilføj Drawing status
            df = self._add_drawing_status(df)
            logging.info("Tilføjet Drawing status")
            
            # Opret Partlist
            partlist_df = self._create_partlist_sheet(df)
            logging.info("Oprettet Partlist sheet")
            
            # Opret kategori sheets
            category_sheets = self._create_category_sheets(df)
            logging.info("Oprettet kategori sheets")
            
            # Opret Compare sheet hvis previous_file er angivet
            compare_df = None
            if self.previous_file and self.previous_file.exists():
                logging.info(f"Indlæser previous BOM: {self.previous_file}")
                previous_df = pd.read_excel(self.previous_file, engine='openpyxl')
                previous_partlist = self._create_partlist_sheet(previous_df)
                compare_df = self._create_compare_sheet(partlist_df, previous_partlist)
                logging.info("Oprettet Compare sheet")
            
            # Send resultat til queue
            self.result = ProcessingResult(success=True, data={
                'dataframe': df,
                'partlist_df': partlist_df,
                'category_sheets': category_sheets,
                'compare_df': compare_df,
                'compare_formatting': self.compare_formatting,
                'original_columns': original_columns,
                'drawings_dir': output_dir
            })
            self.output_queue.put(self.result)
            
            # Anvend formatering på Excel fil
            self._apply_formatting(df)
            
            logging.info("Data processing completed successfully")
            
        except Exception as e:
            logging.error(f"Fejl i data processing: {str(e)}")
            self.result = ProcessingResult(success=False, error=str(e))
            self.output_queue.put(self.result)

class ImageProcessor(threading.Thread):
    """Tråd til billedebehandling"""
    def __init__(self, input_file: Path, output_file: Path, image_queue: queue.Queue):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.image_queue = image_queue
        self.result = ProcessingResult(success=False)
        self._image_handler = ExcelImageHandler()
        self._temp_dir = Path(tempfile.gettempdir()) / "ExcelCopyBOM"
        self._temp_dir.mkdir(exist_ok=True)
        self._temp_files = []
        self._excel = None
        self._workbook = None
        
    def _init_excel(self):
        """Initialiser Excel COM objekt"""
        logging.info("Initialiserer Excel COM objekt for billede udtrækning")
        pythoncom.CoInitialize()
        self._excel = win32com.client.Dispatch("Excel.Application")
        self._excel.Visible = False
        self._excel.DisplayAlerts = False
        
    def _cleanup_excel(self):
        """Luk Excel ned"""
        if self._excel:
            try:
                if self._workbook:
                    self._workbook.Close(SaveChanges=False)
                self._excel.Quit()
            finally:
                pythoncom.CoUninitialize()
        
    def _cleanup_temp_files(self):
        """Ryd op i midlertidige filer"""
        for temp_file in self._temp_files:
            try:
                temp_file.unlink()
            except Exception as e:
                logging.warning(f"Kunne ikke slette temp fil {temp_file}: {str(e)}")
                
    def _extract_and_save_images(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Udtræk billeder fra DataFrame og gem dem som midlertidige filer
        Returns:
            Dict med part number -> temp file sti
        """
        image_mapping = {}
        
        try:
            # Initialiser Excel
            self._init_excel()
            
            # Åbn workbook
            logging.info(f"Åbner workbook for billede udtrækning: {self.input_file}")
            self._workbook = self._excel.Workbooks.Open(str(self.input_file))
            sheet = self._workbook.Sheets(1)
            
            # Find Thumbnail kolonne
            if 'Thumbnail' not in df.columns:
                raise ValueError("Kunne ikke finde Thumbnail kolonne i data")
                
            # Find kolonne index
            thumbnail_col = None
            for col in range(1, sheet.UsedRange.Columns.Count + 1):
                if sheet.Cells(1, col).Value == "Thumbnail":
                    thumbnail_col = col
                    break
                    
            if not thumbnail_col:
                raise ValueError("Kunne ikke finde Thumbnail kolonne i Excel")
            
            # Gennemgå hver række og gem billeder
            for idx, row in df.iterrows():
                if pd.notna(row['Thumbnail']) and pd.notna(row['Part Number']):
                    # Find række i Excel (idx + 2 fordi pandas index starter fra 0 og vi har header)
                    excel_row = idx + 2
                    
                    # Check om der er et billede i cellen
                    cell = sheet.Cells(excel_row, thumbnail_col)
                    shape = None
                    
                    # Find shape i cellen
                    for s in sheet.Shapes:
                        if (s.Left >= cell.Left and 
                            s.Left <= cell.Left + cell.Width and
                            s.Top >= cell.Top and 
                            s.Top <= cell.Top + cell.Height):
                            shape = s
                            break
                    
                    if shape:
                        try:
                            # Gem billede til temp fil
                            temp_file = self._temp_dir / f"{row['Part Number']}.png"
                            self._temp_files.append(temp_file)
                            
                            # Kopier til clipboard
                            shape.Copy()
                            
                            # Gem fra clipboard som PNG
                            image = ImageGrab.grabclipboard()
                            if image:
                                image.save(temp_file, 'PNG')
                                image_mapping[str(row['Part Number'])] = str(temp_file)
                                logging.info(f"Gemt billede for {row['Part Number']} til {temp_file}")
                            
                        except Exception as e:
                            logging.warning(f"Kunne ikke gemme billede for {row['Part Number']}: {str(e)}")
                            continue
                            
        finally:
            self._cleanup_excel()
                
        return image_mapping
        
    def run(self):
        try:
            logging.info("Starter billedebehandling")
            
            # Vent på data fra DataProcessor
            data_result = self.image_queue.get()
            if not data_result.success:
                raise Exception("Data processing fejlede")
                
            # Udtræk DataFrame fra data
            if 'dataframe' not in data_result.data:
                raise ValueError("Mangler DataFrame i processed data")
            
            df = data_result.data['dataframe']
            
            # Udtræk og gem billeder
            logging.info("Udtrækker billeder fra Excel")
            image_mapping = self._extract_and_save_images(df)
            
            if not image_mapping:
                logging.warning("Ingen billeder fundet i Excel filen")
            else:
                logging.info(f"Fandt {len(image_mapping)} billeder")
            
            # Opret ny Excel fil med billeder
            logging.info("Indsætter billeder i ny Excel fil")
            self._image_handler.process_images(
                str(self.output_file),
                image_mapping
            )
            
            self.result = ProcessingResult(success=True)
            
        except Exception as e:
            logging.error(f"Fejl i billedebehandling: {str(e)}")
            self.result = ProcessingResult(success=False, error=str(e))
            
        finally:
            # Ryd op i temp filer
            self._cleanup_temp_files()

class FileProcessor(threading.Thread):
    """Tråd til filkopiering"""
    def __init__(self, source_files: List[Path], target_dir: Path, categories_map: Dict[str, str]):
        super().__init__()
        self.source_files = source_files
        self.target_dir = target_dir
        self.categories_map = categories_map  # Dict[part_number, category]
        self.result = ProcessingResult(success=False)
        
    def _create_category_folders(self) -> Dict[str, Path]:
        """Opret mapper for hver kategori"""
        category_paths = {}
        unique_categories = set(self.categories_map.values())
        
        for category in unique_categories:
            category_path = self.target_dir / category
            category_path.mkdir(parents=True, exist_ok=True)
            category_paths[category] = category_path
            
        return category_paths
        
    def _copy_file(self, source: Path, target: Path) -> bool:
        """Kopier fil med verificering"""
        try:
            import shutil
            shutil.copy2(source, target)
            
            # Verificer at filen er kopieret korrekt
            if not target.exists():
                logging.error(f"Fil blev ikke kopieret: {target}")
                return False
                
            if target.stat().st_size != source.stat().st_size:
                logging.error(f"Fil størrelse matcher ikke for: {target}")
                return False
                
            return True
            
        except Exception as e:
            logging.error(f"Fejl ved kopiering af {source} til {target}: {str(e)}")
            return False
            
    def run(self):
        try:
            logging.info("Starter filkopiering")
            
            # Opret kategori-mapper
            category_paths = self._create_category_folders()
            logging.info(f"Oprettet {len(category_paths)} kategori-mapper")
            
            # Hold styr på kopierede filer
            copied_files = []
            failed_files = []
            
            # Kopier hver fil til den rigtige kategori-mappe
            for source_file in self.source_files:
                # Find part number fra filnavn
                file_part_number = source_file.stem.split("--")[0].strip()
                
                # Find kategori for part number
                if file_part_number in self.categories_map:
                    category = self.categories_map[file_part_number]
                    target_dir = category_paths[category]
                    
                    # Opret target path
                    target_path = target_dir / source_file.name
                    
                    # Kopier filen
                    if self._copy_file(source_file, target_path):
                        copied_files.append(target_path)
                        logging.info(f"Kopieret {source_file.name} til {category}")
                    else:
                        failed_files.append(source_file)
                else:
                    logging.warning(f"Ingen kategori fundet for {file_part_number}")
                    failed_files.append(source_file)
            
            # Log resultater
            logging.info(f"Kopieret {len(copied_files)} filer")
            if failed_files:
                logging.warning(f"Fejl ved kopiering af {len(failed_files)} filer")
                
            self.result = ProcessingResult(
                success=len(failed_files) == 0,
                error=f"Fejl ved kopiering af {len(failed_files)} filer" if failed_files else None,
                data={
                    'copied_files': copied_files,
                    'failed_files': failed_files
                }
            )
            
        except Exception as e:
            logging.error(f"Fejl i filkopiering: {str(e)}")
            self.result = ProcessingResult(success=False, error=str(e))

class ExcelOutlineHandler(threading.Thread):
    """Klasse til at håndtere Excel gruppering og outline"""
    def __init__(self, input_file: Path, output_queue: queue.Queue):
        super().__init__()
        self.input_file = input_file
        self.output_queue = output_queue
        self.result = ProcessingResult(success=False)
        self._excel = None
        self._workbook = None
        
    def _init_excel(self):
        """Initialiser Excel COM objekt"""
        logging.info("Initialiserer Excel COM objekt for outline")
        pythoncom.CoInitialize()
        self._excel = win32com.client.Dispatch("Excel.Application")
        self._excel.Visible = False
        self._excel.DisplayAlerts = False
        
    def _cleanup_excel(self):
        """Luk Excel ned"""
        if self._excel:
            try:
                if self._workbook:
                    self._workbook.Close(SaveChanges=True)
                self._excel.Quit()
            finally:
                pythoncom.CoUninitialize()
                
    def _group_rows(self, df: pd.DataFrame):
        """Gruppér rækker baseret på BOM Structure"""
        try:
            # Åbn workbook
            self._workbook = self._excel.Workbooks.Open(str(self.input_file))
            sheet = self._workbook.Sheets(1)
            
            # Find BOM Structure kolonne
            bom_col = None
            for col in range(1, sheet.UsedRange.Columns.Count + 1):
                if sheet.Cells(1, col).Value == "BOM Structure":
                    bom_col = col
                    break
                    
            if not bom_col:
                raise ValueError("Kunne ikke finde BOM Structure kolonne")
                
            # Hold styr på grupper og niveauer
            current_groups = {}  # niveau -> start_row
            
            # Gennemgå hver række
            for idx, row in df.iterrows():
                excel_row = idx + 2  # +2 for header og 1-baseret index
                
                if pd.notna(row['BOM Structure']):
                    structure = str(row['BOM Structure'])
                    level = len(structure.split('.'))
                    
                    # Afslut tidligere grupper på samme eller højere niveau
                    levels_to_close = [l for l in current_groups.keys() if l >= level]
                    for l in sorted(levels_to_close, reverse=True):
                        start_row = current_groups[l]
                        if excel_row - start_row > 1:  # Kun gruppér hvis der er mere end én række
                            sheet.Rows(f"{start_row}:{excel_row-1}").Group()
                        del current_groups[l]
                    
                    # Start ny gruppe
                    if '.' in structure:  # Kun gruppér under-niveauer
                        current_groups[level] = excel_row
                        
            # Afslut alle åbne grupper
            last_row = len(df) + 1
            for level, start_row in sorted(current_groups.items(), reverse=True):
                if last_row - start_row > 1:
                    sheet.Rows(f"{start_row}:{last_row}").Group()
                    
            # Indstil outline visning
            sheet.Outline.SummaryRow = 1  # 1 = Above, 2 = Below
            sheet.Outline.ShowLevels(RowLevels=1)  # Kollaps til niveau 1
            
            # Gem ændringer
            self._workbook.Save()
            
        except Exception as e:
            raise Exception(f"Fejl ved gruppering af rækker: {str(e)}")
            
    def run(self):
        try:
            logging.info("Starter Excel outline processing")
            
            # Vent på data fra ImageProcessor
            data_result = self.output_queue.get()
            if not data_result.success:
                raise Exception("Image processing fejlede")
                
            # Udtræk DataFrame
            if 'dataframe' not in data_result.data:
                raise ValueError("Mangler DataFrame i processed data")
                
            df = data_result.data['dataframe']
            
            # Initialiser Excel
            self._init_excel()
            
            # Gruppér rækker
            self._group_rows(df)
            
            logging.info("Excel outline processing gennemført")
            self.result = ProcessingResult(success=True)
            
        except Exception as e:
            logging.error(f"Fejl i Excel outline processing: {str(e)}")
            self.result = ProcessingResult(success=False, error=str(e))
            
        finally:
            self._cleanup_excel()

class ThreadedExcelHandler:
    """Hovedklasse til håndtering af parallel Excel processing"""
    def __init__(self, input_file: Path, output_dir: Path):
        self.input_file = input_file
        self.output_dir = output_dir
        self.output_file = output_dir / f"{input_file.stem}_Processed.xlsx"
        
        # Opret queues til kommunikation mellem tråde
        self.data_queue = queue.Queue()
        self.image_queue = queue.Queue()
        self.outline_queue = queue.Queue()
        
        # Initialiser processors
        self.data_processor = DataProcessor(input_file, self.data_queue)
        self.image_processor = ImageProcessor(input_file, self.output_file, self.image_queue)
        self.outline_processor = ExcelOutlineHandler(self.output_file, self.outline_queue)
        self.file_processor = None  # Initialiseres når vi kender filerne
        
    def process(self) -> bool:
        """Start parallel processing af Excel filen"""
        try:
            logging.info(f"Starter threaded processing af {self.input_file}")
            
            # Start data processing
            self.data_processor.start()
            
            # Vent på data processing og send resultat til image processor
            data_result = self.data_queue.get()
            if not data_result.success:
                raise Exception(f"Data processing fejlede: {data_result.error}")
            
            self.image_queue.put(data_result)
            
            # Start image processing
            self.image_processor.start()
            
            # Vent på image processing og send resultat til outline processor
            self.image_processor.join()
            if not self.image_processor.result.success:
                raise Exception(f"Billedebehandling fejlede: {self.image_processor.result.error}")
                
            self.outline_queue.put(data_result)  # Send original data til outline processor
            
            # Start outline processing
            self.outline_processor.start()
            
            # TODO: Start file processor når vi har listen af filer
            
            # Vent på at alle tråde er færdige
            self.data_processor.join()
            self.outline_processor.join()
            if self.file_processor:
                self.file_processor.join()
            
            # Check resultater
            if not self.outline_processor.result.success:
                raise Exception(f"Excel outline processing fejlede: {self.outline_processor.result.error}")
            
            if self.file_processor and not self.file_processor.result.success:
                raise Exception(f"Filkopiering fejlede: {self.file_processor.result.error}")
            
            logging.info("Threaded processing gennemført succesfuldt")
            return True
            
        except Exception as e:
            logging.error(f"Fejl i threaded processing: {str(e)}")
            return False