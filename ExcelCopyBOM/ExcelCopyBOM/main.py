"""
Main entry point for ExcelCopyBOM
"""
import logging
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import config
from gui import ExcelCopyBOMGUI
from ExcelCopyBOM.database import DrawingDatabase
from ExcelCopyBOM.excel_handler import ExcelHandler

def setup_logging():
    """Setup logging til fil og console"""
    try:
        temp_dir = Path(tempfile.gettempdir())
        log_file = temp_dir / "ExcelCopyBOM.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        logging.info(f"Log file location: {log_file}")
        return log_file
    except Exception as e:
        print(f"Failed to setup logging: {str(e)}")
        raise

def process_bom(gui: ExcelCopyBOMGUI):
    """Håndter BOM processing med progress updates"""
    try:
        # Valider input paths
        if not gui.bom_path.get():
            raise ValueError("No BOM file selected")
        
        if not Path(gui.bom_path.get()).exists():
            raise FileNotFoundError(f"BOM file not found: {gui.bom_path.get()}")
        
        # Initialiser Excel handler
        excel = ExcelHandler(
            gui.bom_path.get(),
            gui.prev_bom_path.get() if gui.prev_bom_path.get() else None
        )
        
        # Load workbook
        gui.update_progress(10, "Loading Excel workbook...")
        excel.load_workbook()
        
        # Opdater database hvis valgt
        if gui.update_index.get():
            gui.update_progress(20, "Updating drawing index...")
            db = DrawingDatabase()
            if not db.update_index():
                raise Exception("Failed to update drawing index")
                
        # Process BOM
        gui.update_progress(30, "Processing BOM sheet...")
        excel.process_bom()
        
        # Gem resultat
        output_path = excel.bom_path.parent / f"{excel.bom_path.stem}_Processed.xlsx"
        gui.update_progress(90, "Saving workbook...")
        excel.save_workbook(str(output_path))
        
        # Vis færdig dialog
        gui.update_progress(100, "Complete!")
        gui.show_completion_dialog(str(output_path))
        
    except Exception as e:
        logging.exception("Error during processing")
        gui.show_error(str(e))
        gui.start_button.config(state='normal')

def cleanup(root):
    """Clean up resources before exit"""
    try:
        logging.info("Cleaning up resources...")
        root.destroy()
    except Exception as e:
        logging.error(f"Error during cleanup: {str(e)}")

def cleanup_temp_files():
    """Clean up temporary files"""
    try:
        temp_dir = Path(tempfile.gettempdir())
        for temp_file in temp_dir.glob("ExcelCopyBOM_*"):
            try:
                temp_file.unlink()
                logging.info(f"Deleted temp file: {temp_file}")
            except Exception as e:
                logging.error(f"Failed to delete temp file {temp_file}: {str(e)}")
    except Exception as e:
        logging.error(f"Error cleaning temp files: {str(e)}")

def main():
    """Start ExcelCopyBOM programmet"""
    log_file = setup_logging()
    
    try:
        root = tk.Tk()
        app = ExcelCopyBOMGUI(root)
        root.protocol("WM_DELETE_WINDOW", lambda: cleanup(root))
        root.mainloop()
    except Exception as e:
        logging.exception("Fatal error in main")
        messagebox.showerror("Error", f"Fatal error: {str(e)}")
    finally:
        cleanup_temp_files()

if __name__ == "__main__":
    main()