"""
Test af TRIN 2: Data Indlæsning og Validering
Specifikt test af supplier parts og bogstavsrækker
"""

import logging
import os
import win32com.client
from steps.step2_data import ExcelDataLoader
import time
import shutil
from unittest.mock import patch
from tkinter import messagebox

def create_template_files():
    """Opretter template filer til brug i tests"""
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger('Test')
    excel = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Opret templates for hver test type
        for test_type in ['supplier', 'structure', 'hierarchy']:
            # Opret ny workbook
            wb = excel.Workbooks.Add()
            ws = wb.ActiveSheet
            
            # Tilføj headers
            headers = ['Item', 'Part Number', 'REV', 'BOM Structure', 'Description', 'QTY', 'D', 't', 'L']
            for col, header in enumerate(headers, 1):
                ws.Cells(1, col).Value = header
                
            if test_type == 'supplier':
                test_data = [
                    ['1', '0000-700-001', 'A', 'Part', 'Test Supplier 1', '1', '1', '1', '1'],
                    ['2', '0000-701-002', 'B', 'Part', 'Test Supplier 2', '1', '1', '1', '1'],
                    ['3', '0000-702-003', 'C', 'Part', 'Test Supplier 3', '1', '1', '1', '1'],
                    ['4', '4003-01-A01', 'A', 'Part', 'Normal Part', '1', '1', '1', '1'],
                    ['5', 'ABCD', '', 'Part', 'Bogstaver 1', '1', '1', '1', '1'],
                    ['6', 'XYZ', '', 'Part', 'Bogstaver 2', '1', '1', '1', '1']
                ]
            elif test_type == 'structure':
                test_data = [
                    ['1', '4003-01-A01', 'A', 'Inseparable', 'Top Assembly', '1', '1', '1', '1'],
                    ['1.1', '4003-02-B01', 'A', '|Part', 'Should be deleted', '2', '1', '1', '1'],
                    ['1.2', '4003-02-B02', 'A', '|Part', 'Should be deleted', '1', '1', '1', '1'],
                    ['2', '0000-301-001', 'A', 'Assembly', 'Basic Component', '1', '1', '1', '1'],
                    ['2.1', '4003-02-B03', 'A', '|Part', 'Should be deleted', '1', '1', '1', '1'],
                    ['3', '4003-01-A02', 'A', 'Assembly', 'Normal Assembly', '1', '1', '1', '1'],
                    ['3.1', '4003-02-B04', 'A', '|Phantom', 'Should be deleted', '2', '1', '1', '1'],
                    ['3.2', '4003-02-B05', 'A', '|Part', 'Should remain', '1', '1', '1', '1']
                ]
            else:  # hierarchy
                test_data = [
                    ['1', '4003-01-A01', 'A', 'Assembly', 'Top Assembly', '1', '1', '1', '1'],
                    ['1.1', '4003-02-B01', 'A', '|Part', 'Sub Part 1', '2', '1', '1', '1'],
                    ['1.2', '4003-02-B02', 'A', '|Assembly', 'Sub Assembly 1', '1', '1', '1', '1'],
                    ['1.2.1', '4003-03-C01', 'A', '||Part', 'Sub Sub Part 1', '3', '1', '1', '1'],
                    ['1.2.2', '4003-03-C02', 'A', '||Part', 'Sub Sub Part 2', '1', '1', '1', '1'],
                    ['1.3', '4003-02-B03', 'A', '|Part', 'Sub Part 2', '1', '1', '1', '1'],
                    ['2', '4003-01-A02', 'A', 'Assembly', 'Top Assembly 2', '1', '1', '1', '1'],
                    ['2.1', '4003-02-B04', 'A', '|Part', 'Sub Part 3', '2', '1', '1', '1']
                ]
                
            # Indsæt test data
            for row, data in enumerate(test_data, 2):
                for col, value in enumerate(data, 1):
                    ws.Cells(row, col).Value = value
                    
            # Gem template fil
            template_file = os.path.join(os.path.dirname(__file__), f'template_bom_{test_type}.xlsx')
            wb.SaveAs(template_file)
            wb.Close()
            logger.info(f"Template fil oprettet: {template_file}")
            
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass

def setup_test_file(test_type='supplier'):
    """Opretter en test Excel fil ved at kopiere TestOrg.xlsx"""
    # Brug TestOrg.xlsx som template
    template_file = os.path.join(os.path.dirname(__file__), 'test', 'TestOrg.xlsx')
    
    # Opret test fil med korrekt navn
    if test_type == 'supplier':
        test_file = os.path.join(os.path.dirname(__file__), '4003-02.1-A01-- - BOM.xlsx')
    elif test_type == 'hierarchy':
        test_file = os.path.join(os.path.dirname(__file__), '4003-02.1-A01-A - BOM.xlsx')
    else:  # structure
        test_file = os.path.join(os.path.dirname(__file__), '4003-02-A01-B - BOM.xlsx')
    
    # Kopier template til test fil
    shutil.copy2(template_file, test_file)
    return test_file

def test_supplier_parts():
    """Tester håndtering af supplier parts"""
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger('Test')
    
    try:
        # Opret test fil
        test_file = setup_test_file()
        logger.info(f"Test fil oprettet: {test_file}")
        
        # Test med include_suppliers = False
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(test_file)
        
        try:
            logger.info("Tester uden supplier parts...")
            loader = ExcelDataLoader(wb, include_suppliers=False)
            loader.process_file()
            
            # Tjek antal rækker (skal være 3 mindre end oprindeligt)
            final_rows = wb.ActiveSheet.UsedRange.Rows.Count
            logger.info(f"Antal rækker efter behandling: {final_rows}")
            
            # Tjek om supplier parts er fjernet
            part_numbers = []
            for row in range(2, final_rows + 1):
                part_number = str(wb.ActiveSheet.Cells(row, 2).Value).strip()
                part_numbers.append(part_number)
                
            logger.info(f"Resterende part numbers: {part_numbers}")
            
            # Der bør ikke være nogen supplier parts tilbage
            supplier_parts = [pn for pn in part_numbers if pn.startswith(('0000-700-', '0000-701-', '0000-702-'))]
            if supplier_parts:
                logger.error(f"Fandt supplier parts der ikke skulle være der: {supplier_parts}")
            else:
                logger.info("Alle supplier parts blev fjernet korrekt")
                
            # Gem en kopi i test-mappen
            test_copy = os.path.join(os.path.dirname(__file__), 'test', 'test_bom_supplier_result.xlsx')
            wb.SaveAs(test_copy)
            logger.info(f"Test resultat gemt som: {test_copy}")
                
        finally:
            wb.Close(SaveChanges=False)
            excel.Quit()
            
        # Opryd original test fil
        os.remove(test_file)
        logger.info("Original test fil fjernet")
        
    except Exception as e:
        logger.error(f"Fejl under test: {str(e)}", exc_info=True)

def test_bom_hierarchy():
    """Tester BOM hierarki nummerering og gruppering"""
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger('Test')
    
    try:
        # Opret test fil
        test_file = setup_test_file('hierarchy')
        logger.info(f"Test fil oprettet: {test_file}")
        
        # Test hierarki
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Sæt til True så vi kan se grupperingen
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(test_file)
        
        try:
            logger.info("Tester BOM hierarki...")
            # Mock messagebox.askyesno til at returnere False
            with patch('tkinter.messagebox.askyesno', return_value=False):
                loader = ExcelDataLoader(wb, include_suppliers=False)
                loader.process_file()
            
            # Tjek item numre
            item_col = loader.columns['Item']
            last_row = wb.ActiveSheet.UsedRange.Rows.Count
            actual_items = []
            
            for row in range(2, last_row + 1):
                item = str(wb.ActiveSheet.Cells(row, item_col).Value).strip()
                if item:
                    actual_items.append(item)
            
            logger.info(f"Item numre efter behandling: {actual_items}")
            
            # Tjek at hierarkiet er korrekt
            for i in range(len(actual_items) - 1):
                current = actual_items[i]
                next_item = actual_items[i + 1]
                
                # Tjek at niveauerne er korrekte
                current_parts = current.split('.')
                next_parts = next_item.split('.')
                
                # Næste item skal enten:
                # 1. Være et child (have et niveau mere)
                # 2. Være en sibling (samme antal niveauer)
                # 3. Være en ny parent (færre niveauer)
                level_diff = len(next_parts) - len(current_parts)
                if level_diff > 1:
                    logger.error(f"For stort spring i hierarki mellem {current} og {next_item}")
                    return
                
                # Hvis det er et child, tjek at parent delen matcher
                if level_diff == 1:
                    if '.'.join(next_parts[:-1]) != current:
                        logger.error(f"Child {next_item} matcher ikke parent {current}")
                        return
            
            logger.info("BOM hierarki er korrekt")
            
            # Gem resultat for manuel verifikation af gruppering
            result_file = os.path.join(os.path.dirname(__file__), 'test', 'test_bom_hierarchy_result.xlsx')
            wb.SaveAs(result_file)
            logger.info(f"Test resultat gemt som: {result_file}")
            
        finally:
            wb.Close(SaveChanges=False)
            
    except Exception as e:
        logger.error(f"Fejl under test: {str(e)}", exc_info=True)
        
    finally:
        try:
            excel.Quit()
        except:
            pass
        
        # Fjern test fil
        if os.path.exists(test_file):
            os.remove(test_file)
            logger.info("Original test fil fjernet")

def test_bom_structure():
    """Tester håndtering af BOM Structure regler"""
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger('Test')
    excel = None
    
    try:
        # Opret test fil
        test_file = setup_test_file('structure')
        logger.info(f"Test fil oprettet: {test_file}")
        
        # Test BOM Structure regler
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Sæt til True så vi kan se resultatet
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(test_file)
        
        logger.info("Tester BOM Structure regler...")
        # Mock messagebox.askyesno til at returnere False
        with patch('tkinter.messagebox.askyesno', return_value=False):
            loader = ExcelDataLoader(wb, include_suppliers=False)  # Tilføj include_suppliers=False
            loader.process_file()
        
        # Tjek resultatet
        remaining_rows = []
        last_row = wb.ActiveSheet.UsedRange.Rows.Count
        for row in range(2, last_row + 1):
            item = str(wb.ActiveSheet.Cells(row, 1).Value).strip()
            part_number = str(wb.ActiveSheet.Cells(row, 2).Value).strip()
            structure = str(wb.ActiveSheet.Cells(row, 4).Value).strip()
            remaining_rows.append((item, part_number, structure))
        
        logger.info("Resterende rækker:")
        for row in remaining_rows:
            logger.info(f"Item: {row[0]}, Part: {row[1]}, Structure: {row[2]}")
        
        # Gem en kopi i test-mappen
        test_copy = os.path.join(os.path.dirname(__file__), 'test', 'test_bom_structure_result.xlsx')
        wb.SaveAs(test_copy)
        logger.info(f"Test resultat gemt som: {test_copy}")
        
    except Exception as e:
        logger.error(f"Fejl under test: {str(e)}", exc_info=True)
        
    finally:
        if excel:
            try:
                wb.Close(SaveChanges=False)
                excel.Quit()
            except:
                pass
            
        # Opryd original test fil
        try:
            os.remove(test_file)
            logger.info("Original test fil fjernet")
        except:
            pass

if __name__ == "__main__":
    # Opret template filer hvis de ikke findes
    for test_type in ['supplier', 'structure', 'hierarchy']:
        template_file = os.path.join(os.path.dirname(__file__), f'template_bom_{test_type}.xlsx')
        if not os.path.exists(template_file):
            create_template_files()
            break
    
    # Kør tests
    test_supplier_parts()
    time.sleep(1)
    
    test_bom_hierarchy()
    time.sleep(1)
    
    test_bom_structure() 