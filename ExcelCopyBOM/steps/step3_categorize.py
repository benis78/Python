"""
# =============================================
# LOCKED VERSION - DO NOT MODIFY WITHOUT TESTS
# =============================================
# Denne fil håndterer kategorisering af part numbers baseret på deres struktur.
# Version: 1.0.0
# Låst: [CURRENT_DATE]
# 
# Kategorisering følger disse regler:
# 1. Basic Components (0000-3xx-xxx)
# 2. Suppliers Parts (0000-7xx-xxx)
# 3. Project Specific Parts (xxxx-xx-PSxx)
# 4. Piping (præfikser: BM, AF, AA, BG, CD, etc.)
# 5. Equipment (6xx)
# 6. Area Drawings (A=Arrangement, E=Equipment, etc.)
# 7. Layout Drawings
#
# Alle part numbers der ikke matcher disse regler kategoriseres som "Other Parts"
"""

import logging
import os
import re
from typing import Tuple, Dict
import yaml

class PartNumberCategorizer:
    def __init__(self):
        self.logger = logging.getLogger('ExcelCopyBOM.Categorizer')
        self.categories = self._load_categories()
        
    def _load_categories(self) -> dict:
        """Indlæser kategorier fra Categories.yaml"""
        yaml_path = os.path.join(os.path.dirname(__file__), "..", "Categories.yaml")
        
        try:
            with open(yaml_path, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                self.logger.debug(f"Indlæst kategorier fra {yaml_path}")
                return data['categories']
                
        except Exception as e:
            self.logger.error(f"Fejl ved indlæsning af kategorier: {str(e)}")
            return {}

    def _parse_part_number(self, part_number: str) -> Dict[str, str]:
        """
        Parser et part number og returnerer dets komponenter.
        F.eks. "4003-04.1-BM004-01-01" bliver til:
        {
            'identifier1': '4003',
            'identifier2': '04.1',
            'identifier3': 'BM004',
            'full_number': '4003-04.1-BM004-01-01'
        }
        """
        if not part_number or '-' not in part_number:
            return {}
            
        parts = part_number.split('-')
        if len(parts) < 3:
            return {}
            
        # Identifier 1 skal være 0000 eller >= 2000
        identifier1 = parts[0]
        if not identifier1.isdigit() or (identifier1 != "0000" and int(identifier1) < 2000):
            return {}
            
        # Identifier 2 kan være ## eller ##.# eller ###
        identifier2 = parts[1]
        if not (re.match(r'^\d{2}$', identifier2) or 
                re.match(r'^\d{2}\.\d$', identifier2) or 
                re.match(r'^\d{3}$', identifier2)):
            return {}
            
        # Find identifier3 i gruppe 3, 4 eller 5
        identifier3 = None
        for part in parts[2:5]:
            # Tjek for gyldige identifier3 formater
            if (re.match(r'^\d{3}$', part) or  # ###
                re.match(r'^[A-Z]\d{2}$', part) or  # A##
                re.match(r'^[A-Z]{2}\d+', part) or  # AA###
                part.startswith(('PS', 'A', 'B', 'E', 'F', 'BM', 'AF', 'CD', 'HW'))):
                identifier3 = part
                break
                
        if not identifier3:
            return {}
            
        return {
            'identifier1': identifier1,
            'identifier2': identifier2,
            'identifier3': identifier3,
            'full_number': part_number
        }

    def categorize(self, part_number: str) -> Tuple[str, str]:
        """
        Kategoriserer et part number og returnerer (type, category).
        Hvis part number ikke kan kategoriseres, returneres ("Other Parts", "Other Parts").
        """
        # Parser part number
        parsed = self._parse_part_number(part_number)
        if not parsed:
            return "Unknown", "Unknown"
            
        # Tjek Basic Components (0000-3xx)
        if parsed['identifier1'] == "0000" and parsed['identifier2'].startswith("3"):
            category_data = self.categories.get('basic_components')
            if category_data:
                identifier2_num = str(int(parsed['identifier2']))
                if identifier2_num in category_data['types']:
                    return (category_data['types'][identifier2_num], 
                           category_data['category'])
                           
        # Tjek Suppliers Parts (0000-7xx)
        if parsed['identifier1'] == "0000" and parsed['identifier2'].startswith("7"):
            category_data = self.categories.get('suppliers_parts')
            if category_data:
                identifier2_num = str(int(parsed['identifier2']))
                if identifier2_num in category_data['types']:
                    return (category_data['types'][identifier2_num], 
                           category_data['category'])
                           
        # Tjek Project Specific Parts (PS)
        if parsed['identifier3'].startswith('PS'):
            return "Project Specific Parts", "Area Drawings"
            
        # Tjek Piping
        category_data = self.categories.get('piping')
        if category_data:
            for prefix in category_data['identifiers']['identifier3']:
                if parsed['identifier3'].startswith(prefix):
                    return category_data['types'][prefix], category_data['category']
                    
        # Tjek Equipment numbers (6xx)
        if parsed['identifier2'].startswith("6") and len(parsed['identifier2']) == 3:
            category_data = self.categories.get('equipment')
            if category_data:
                identifier2_num = str(int(parsed['identifier2']))
                if identifier2_num in category_data['types']:
                    # Tjek om identifier3 er gyldig for equipment
                    first_char = parsed['identifier3'][0]
                    if first_char in category_data['identifiers']['identifier3']:
                        return (category_data['types'][identifier2_num], 
                               category_data['categories'][identifier2_num])
                               
        # Tjek Area Drawings
        category_data = self.categories.get('area_drawings')
        if category_data:
            # Kun tjek area drawings hvis identifier2 IKKE er et equipment number
            if not (parsed['identifier2'].startswith("6") and len(parsed['identifier2']) == 3):
                first_char = parsed['identifier3'][0]
                if first_char in category_data['types']:
                    return category_data['types'][first_char], category_data['category']
                    
        # Tjek Layout Drawings
        category_data = self.categories.get('layout_drawings')
        if category_data and parsed['identifier3'].isdigit():
            identifier3_num = str(int(parsed['identifier3']))
            if identifier3_num in category_data['types']:
                return category_data['types'][identifier3_num], category_data['category']
                
        return "Unknown", "Unknown" 