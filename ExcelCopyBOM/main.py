"""Main entry point for ExcelCopyBOM"""

import sys
from pathlib import Path
from PyQt5.QtWidgets import QApplication
from gui import MainWindow

def main():
    """Start ExcelCopyBOM programmet"""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 