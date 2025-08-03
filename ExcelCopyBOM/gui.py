"""GUI for ExcelCopyBOM"""

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

# Opsæt logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='excelcopybom.log'
)

class MainWindow(QMainWindow):
    """Hovedvindue for ExcelCopyBOM"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ExcelCopyBOM")
        self.setGeometry(100, 100, 800, 400)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)  # Hold vinduet øverst
        
        logging.info("Starter ExcelCopyBOM")
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Open Excel BOM List
        bom_layout = QHBoxLayout()
        self.bom_label = QLabel("Open Excel BOM List:")
        self.bom_path = QLabel("No file selected")
        self.bom_browse = QPushButton("Browse")
        self.bom_browse.clicked.connect(self.select_bom_file)
        bom_layout.addWidget(self.bom_label)
        bom_layout.addWidget(self.bom_path)
        bom_layout.addWidget(self.bom_browse)
        layout.addLayout(bom_layout)
        
        # Previous BOM List
        prev_layout = QHBoxLayout()
        self.prev_label = QLabel("Previous BOM List:")
        self.prev_path = QLabel("No file selected")
        self.prev_browse = QPushButton("Browse")
        self.prev_browse.clicked.connect(self.select_prev_file)
        prev_layout.addWidget(self.prev_label)
        prev_layout.addWidget(self.prev_path)
        prev_layout.addWidget(self.prev_browse)
        layout.addLayout(prev_layout)
        
        # Update Index File checkbox
        self.update_index = QCheckBox("Update Index File")
        layout.addWidget(self.update_index)
        
        # Progress section
        self.progress_label = QLabel("Ready")
        layout.addWidget(self.progress_label)
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # Add stretch to push Start button to bottom
        layout.addStretch()
        
        # Start button at bottom
        self.start_btn = QPushButton("Start")
        self.start_btn.clicked.connect(self.start_processing)
        self.start_btn.setEnabled(False)
        layout.addWidget(self.start_btn)
        
        # Instance variables
        self.bom_file = None
        self.prev_file = None
        self.processor = None
        self.result = None
        self.result_queue = queue.Queue()
        self.output_folder = None
        
        # Timer til at checke queue
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_queue)
        self.timer.start(100)  # Check hver 100ms
        
    def select_bom_file(self):
        """Vælg Excel BOM fil"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select Excel BOM file",
                "",
                "Excel files (*.xlsx *.xls)"
            )
            
            if file_path:
                self.bom_file = Path(file_path)
                self.bom_path.setText(self.bom_file.name)
                logging.info(f"Valgt BOM fil: {self.bom_file}")
                self.check_start_enabled()
        except Exception as e:
            logging.error(f"Fejl ved valg af BOM fil: {str(e)}")
            self.show_error("Fejl ved valg af BOM fil", str(e))
            
    def select_prev_file(self):
        """Vælg tidligere BOM fil"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select previous BOM file",
                "",
                "Excel files (*.xlsx *.xls)"
            )
            
            if file_path:
                self.prev_file = Path(file_path)
                self.prev_path.setText(self.prev_file.name)
                logging.info(f"Valgt tidligere BOM fil: {self.prev_file}")
                self.check_start_enabled()
        except Exception as e:
            logging.error(f"Fejl ved valg af tidligere BOM fil: {str(e)}")
            self.show_error("Fejl ved valg af tidligere BOM fil", str(e))
            
    def check_start_enabled(self):
        """Aktiver Start knappen hvis en BOM fil er valgt"""
        self.start_btn.setEnabled(bool(self.bom_file))
            
    def check_network_connection(self) -> bool:
        """Tjek om der er forbindelse til netværksdrevet"""
        try:
            network_path = os.path.dirname(config.DRAWING_DB_PATH)
            logging.info(f"Tjekker netværksforbindelse til: {network_path}")
            
            # Tjek om stien eksisterer
            if not os.path.exists(network_path):
                logging.warning(f"Kan ikke få adgang til netværkssti: {network_path}")
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setWindowTitle("Netværksfejl")
                msg.setText("Kan ikke få forbindelse til tegningsdrevet!")
                msg.setInformativeText(
                    f"Kontroller at du har adgang til:\n{network_path}\n\n"
                    "1. Er du på firmanetværket?\n"
                    "2. Er netværksdrevet tilgængeligt?\n"
                    "3. Har du de nødvendige tilladelser?"
                )
                msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Retry)
                
                # Hvis brugeren vælger Retry, prøv igen
                if msg.exec_() == QMessageBox.Retry:
                    logging.info("Bruger valgte at prøve igen")
                    return self.check_network_connection()
                return False
                
            logging.info("Netværksforbindelse OK")
            return True
            
        except Exception as e:
            logging.error(f"Fejl ved tjek af netværksforbindelse: {str(e)}")
            self.show_error("Netværksfejl", str(e))
            return False
            
    def show_error(self, title: str, message: str):
        """Vis fejlmeddelelse"""
        logging.error(f"{title}: {message}")
        QMessageBox.critical(self, title, message)
            
    def start_processing(self):
        """Start behandling af filer"""
        try:
            logging.info("Starter behandling")
            
            # Tjek netværksforbindelse først
            if not self.check_network_connection():
                logging.warning("Behandling afbrudt pga. manglende netværksforbindelse")
                return
                
            # Opdater index hvis valgt
            if self.update_index.isChecked():
                self.progress_label.setText("Updating file index...")
                logging.info("Opdaterer filindex")
                
                if os.path.exists(config.DRAWING_INDEXER_PATH):
                    try:
                        subprocess.run([config.DRAWING_INDEXER_PATH], check=True)
                        logging.info("Filindex opdateret")
                    except subprocess.CalledProcessError as e:
                        logging.error(f"Fejl ved kørsel af file_indexer.exe: {str(e)}")
                        self.show_error(
                            "Indexeringsfejl",
                            f"Fejl ved kørsel af file_indexer.exe:\n{str(e)}"
                        )
                else:
                    logging.warning(f"Kan ikke finde file_indexer.exe: {config.DRAWING_INDEXER_PATH}")
                    QMessageBox.warning(
                        self,
                        "Advarsel",
                        f"Kan ikke finde file_indexer.exe på:\n{config.DRAWING_INDEXER_PATH}\nFortsætter uden at opdatere index."
                    )
            
            # Nulstil GUI
            self.progress_bar.setValue(0)
            self.progress_label.setText("Processing files...")
            self.start_btn.setEnabled(False)
            
            # Start behandling i separat tråd
            logging.info("Starter DataProcessor")
            self.processor = DataProcessor(
                self.bom_file,
                self.result_queue,
                previous_file=self.prev_file
            )
            self.processor.start()
            
        except Exception as e:
            logging.error(f"Fejl ved start af behandling: {str(e)}")
            self.show_error("Behandlingsfejl", str(e))
            
    def check_queue(self):
        """Check queue for resultater og opdater progress"""
        try:
            while True:  # Tøm køen
                result = self.result_queue.get_nowait()
                if result.success:
                    self.result = result.data
                    self.output_folder = result.data.get('output_folder')
                    self.progress_bar.setValue(100)
                    self.progress_label.setText("Processing complete!")
                    logging.info("Behandling færdig")
                    self.show_done_dialog()
                else:
                    logging.error(f"Fejl under behandling: {result.message}")
                    self.show_error("Behandlingsfejl", result.message)
        except queue.Empty:
            pass  # Ingen flere resultater i køen
            
    def show_done_dialog(self):
        """Vis Done dialog og åbn output mappe"""
        msg = QMessageBox()
        msg.setWindowTitle("Done")
        msg.setText("Processing complete!")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.buttonClicked.connect(self.done_clicked)
        msg.exec_()
        
    def done_clicked(self, button):
        """Håndter klik på Done knappen"""
        if self.output_folder:
            # Åbn output mappe
            if os.path.exists(self.output_folder):
                logging.info(f"Åbner output mappe: {self.output_folder}")
                os.startfile(self.output_folder)
            else:
                logging.warning(f"Output mappe findes ikke: {self.output_folder}")
        # Luk programmet
        logging.info("Lukker programmet")
        QApplication.quit() 