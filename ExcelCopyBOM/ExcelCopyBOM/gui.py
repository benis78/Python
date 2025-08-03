"""
GUI for ExcelCopyBOM
"""
import os
import queue
import subprocess
import logging
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QLabel, QProgressBar, QMessageBox, QCheckBox,
    QApplication
)
from PyQt5.QtCore import Qt, QThread, QTimer
import config
from threaded_excel_handler import DataProcessor

class ExcelCopyBOMGUI:
    def __init__(self):
        self.root = QMainWindow()
        self.root.setWindowTitle(config.WINDOW_TITLE)
        self.root.setGeometry(config.WINDOW_SIZE)
        
        # Sæt vinduet til at være øverst ved start
        self.root.setWindowFlags(self.root.windowFlags() | Qt.WindowStaysOnTopHint)
        self.root.show()
        self.root.activateWindow()
        self.root.raise_()
        
        # File paths
        self.bom_path = config.StringVar()
        self.prev_bom_path = config.StringVar()
        self.update_index = config.BooleanVar()
        
        self._create_widgets()
        
    def _create_widgets(self):
        """Opret alle GUI elementer"""
        # Main frame med padding
        main_frame = QWidget()
        layout = QVBoxLayout()
        main_frame.setLayout(layout)
        
        # BOM file selection
        layout.addWidget(QLabel("Open Excel BOM List:"))
        layout.addWidget(QLineEdit(textvariable=self.bom_path, width=50))
        layout.addWidget(QPushButton("Browse", clicked=self._browse_bom))
        
        # Previous BOM file selection
        layout.addWidget(QLabel("Previous BOM List:"))
        layout.addWidget(QLineEdit(textvariable=self.prev_bom_path, width=50))
        layout.addWidget(QPushButton("Browse", clicked=self._browse_prev_bom))
        
        # Update index checkbox
        layout.addWidget(QCheckBox("Update Index File", variable=self.update_index))
        
        # Progress bar og label
        self.progress_var = config.DoubleVar()
        self.progress_label = QLabel()
        layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Start button
        self.start_button = QPushButton("Start Processing")
        self.start_button.clicked.connect(self._start_processing)
        layout.addWidget(self.start_button)
        
    def _browse_bom(self):
        """Åbn file dialog for BOM fil"""
        filename, _ = QFileDialog.getOpenFileName(self.root, "Select BOM File", "", "Excel files (*.xlsx);;All files (*.*)")
        if filename:
            self.bom_path.set(filename)
            
    def _browse_prev_bom(self):
        """Åbn file dialog for tidligere BOM fil"""
        filename, _ = QFileDialog.getOpenFileName(self.root, "Select Previous BOM File", "", "Excel files (*.xlsx);;All files (*.*)")
        if filename:
            self.prev_bom_path.set(filename)
            
    def _start_processing(self):
        """Start behandling af BOM fil"""
        if not self.bom_path.get():
            QMessageBox.warning(self.root, "Warning", "Please select a BOM file")
            return
            
        self.start_button.setEnabled(False)
        # Her skal vi kalde hovedprocessen
        
    def update_progress(self, percentage, message):
        """Opdater progress bar og besked"""
        self.progress_var.set(percentage)
        self.progress_label.setText(message)
        self.root.update()
        
    def show_completion_dialog(self, output_path):
        """Vis færdiggørelses dialog og åbn output mappe"""
        result = QMessageBox.information(self.root, "Complete", "Processing completed successfully!", QMessageBox.Ok)
        if result == QMessageBox.Ok:
            os.startfile(str(Path(output_path).parent))
            self.root.close()
            
    def show_error(self, message, log_path=None):
        """Vis fejlbesked og evt. log fil placering"""
        error_msg = f"An error occurred: {message}"
        if log_path:
            error_msg += f"\nLog file has been saved to: {log_path}"
        QMessageBox.warning(self.root, "Warning", error_msg)
        
    def run(self):
        """Start GUI main loop"""
        self.root.exec_()

def get_drawing_database():
    from ExcelCopyBOM.database import DrawingDatabase
    return DrawingDatabase() 