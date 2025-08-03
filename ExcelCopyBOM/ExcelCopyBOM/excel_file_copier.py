"""
Filkopierings modul til ExcelCopyBOM (TRIN 5)
Håndterer kopiering af PDF og DWG filer samt logging af revisionsforskelle
"""

import os
import shutil
from datetime import datetime
import re
from typing import Dict, List, Set, Tuple
from tkinter import messagebox
import logging

class ExcelFileCopier:
    def __init__(self, source_dir: str, target_dir: str, before_date: datetime = None):
        """
        Initialiserer fil kopieringsmodulet
        :param source_dir: Kildemappen med PDF/DWG filer
        :param target_dir: Målmappen hvor filerne skal kopieres til
        :param before_date: Hvis sat, kun kopier filer fra før denne dato
        """
        self.source_dir = source_dir
        self.target_dir = target_dir
        self.before_date = before_date
        self.rev_differences: Dict[str, Dict[str, str]] = {}
        
        # Opsæt logging
        self.logger = logging.getLogger('ExcelFileCopier')
        self.logger.setLevel(logging.DEBUG)
        
        # Tilføj fil handler hvis den ikke allerede findes
        if not self.logger.handlers:
            log_file = os.path.join(target_dir, 'file_copy.log')
            handler = logging.FileHandler(log_file, encoding='utf-8')
            handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            self.logger.addHandler(handler)
    
    def find_latest_revision(self, part_number: str, file_type: str) -> Tuple[str, datetime]:
        """
        Finder den seneste revision af en fil før before_date
        :param part_number: Part number at søge efter
        :param file_type: Filtype (pdf eller dwg)
        :return: Tuple med (filnavn, fil dato)
        """
        latest_file = ''
        latest_date = datetime.min
        
        # Regex pattern for at matche filnavne
        pattern = rf"{part_number}-[A-Z]\d*\.{file_type}$"
        
        for root, _, files in os.walk(self.source_dir):
            for file in files:
                if re.match(pattern, file, re.IGNORECASE):
                    file_path = os.path.join(root, file)
                    file_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                    
                    # Check om filen er før before_date hvis sat
                    if self.before_date and file_date > self.before_date:
                        continue
                        
                    if file_date > latest_date:
                        latest_file = file
                        latest_date = file_date
        
        return latest_file, latest_date
    
    def get_revision_from_filename(self, filename: str) -> str:
        """
        Udtrækker revision fra filnavn
        :param filename: Filnavn at udtrække fra
        :return: Revisionsbogstav
        """
        match = re.search(r"-([A-Z])\d*\.", filename)
        return match.group(1) if match else ''
    
    def copy_file(self, source_file: str, target_file: str) -> bool:
        """
        Kopierer en fil med fejlhåndtering
        :param source_file: Kildefil
        :param target_file: Målfil
        :return: True hvis kopiering lykkedes
        """
        try:
            # Opret målmappe hvis den ikke findes
            os.makedirs(os.path.dirname(target_file), exist_ok=True)
            
            # Kopier filen
            shutil.copy2(source_file, target_file)
            self.logger.info(f"Kopieret: {os.path.basename(source_file)} -> {os.path.basename(target_file)}")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under kopiering af {source_file}: {str(e)}")
            return False
    
    def process_part_number(self, part_number: str) -> None:
        """
        Behandler et part number og kopierer tilhørende filer
        :param part_number: Part number at behandle
        """
        # Find seneste PDF og DWG filer
        pdf_file, pdf_date = self.find_latest_revision(part_number, "pdf")
        dwg_file, dwg_date = self.find_latest_revision(part_number, "dwg")
        
        if not pdf_file and not dwg_file:
            self.logger.warning(f"Ingen filer fundet for {part_number}")
            return
            
        # Check for revisionsforskelle
        if pdf_file and dwg_file:
            pdf_rev = self.get_revision_from_filename(pdf_file)
            dwg_rev = self.get_revision_from_filename(dwg_file)
            
            if pdf_rev != dwg_rev:
                self.rev_differences[part_number] = {
                    'pdf_rev': pdf_rev,
                    'dwg_rev': dwg_rev,
                    'pdf_date': pdf_date.strftime('%Y-%m-%d %H:%M:%S'),
                    'dwg_date': dwg_date.strftime('%Y-%m-%d %H:%M:%S')
                }
                self.logger.warning(
                    f"Revisionsforskelle for {part_number}: "
                    f"PDF={pdf_rev} ({pdf_date}), DWG={dwg_rev} ({dwg_date})"
                )
        
        # Kopier filer
        if pdf_file:
            source_path = os.path.join(self.source_dir, pdf_file)
            target_path = os.path.join(self.target_dir, pdf_file)
            self.copy_file(source_path, target_path)
            
        if dwg_file:
            source_path = os.path.join(self.source_dir, dwg_file)
            target_path = os.path.join(self.target_dir, dwg_file)
            self.copy_file(source_path, target_path)
    
    def save_revision_report(self) -> None:
        """Gemmer rapport over revisionsforskelle"""
        if not self.rev_differences:
            return
            
        report_path = os.path.join(self.target_dir, 'revision_differences.txt')
        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("Rapport over revisionsforskelle mellem PDF og DWG filer\n")
                f.write("=" * 60 + "\n\n")
                
                for part_number, data in sorted(self.rev_differences.items()):
                    f.write(f"Part Number: {part_number}\n")
                    f.write(f"PDF Revision: {data['pdf_rev']} (fra {data['pdf_date']})\n")
                    f.write(f"DWG Revision: {data['dwg_rev']} (fra {data['dwg_date']})\n")
                    f.write("-" * 40 + "\n")
                    
            self.logger.info(f"Revisionsrapport gemt: {report_path}")
            
        except Exception as e:
            self.logger.error(f"Fejl under gemning af revisionsrapport: {str(e)}")
    
    def process_part_numbers(self, part_numbers: Set[str]) -> bool:
        """
        Hovedfunktion der behandler en liste af part numbers
        :param part_numbers: Set af part numbers der skal behandles
        :return: True hvis succesfuld
        """
        try:
            self.logger.info(f"Starter behandling af {len(part_numbers)} part numbers")
            
            for part_number in sorted(part_numbers):
                self.process_part_number(part_number)
            
            self.save_revision_report()
            self.logger.info("Filkopiering fuldført")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under filkopiering: {str(e)}")
            messagebox.showerror("Error", f"Fejl under filkopiering:\n{str(e)}")
            return False

if __name__ == "__main__":
    # Test kode
    test_source = r"C:\Test\Source"
    test_target = r"C:\Test\Target"
    test_parts = {"1234-01", "1234-02"}
    
    copier = ExcelFileCopier(test_source, test_target)
    copier.process_part_numbers(test_parts) 