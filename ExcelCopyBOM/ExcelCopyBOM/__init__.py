"""
ExcelCopyBOM package
"""

from .threaded_excel_handler import (
    ThreadedExcelHandler,
    DataProcessor,
    ImageProcessor,
    FileProcessor,
    ExcelOutlineHandler,
    ProcessingResult
)

from .gui import ExcelCopyBOMGUI
from .database import DrawingDatabase
from .Categories import PartNumberParser

__version__ = "0.1" 