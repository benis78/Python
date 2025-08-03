import yaml
import os
import re
from typing import Dict, Tuple
import logging

class CategoryParser:
    def __init__(self, yaml_file: str = None):
        """Initialize the CategoryParser with YAML configuration"""
        if yaml_file is None:
            # Først prøv at bruge filen fra netværksdrevet
            network_yaml = r'\\192.168.170.18\drawings\Categories.yaml'
            if os.path.exists(network_yaml):
                yaml_file = network_yaml
            else:
                # Hvis netværksfilen ikke findes, brug den lokale fil
                yaml_file = os.path.join(os.path.dirname(__file__), "Categories.yaml")
        
        with open(yaml_file, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)['categories']

    def parse_part_number(self, part_number: str) -> Dict:
        """
        Parse part number into components following the format rules:
        - Minimum 3 groups separated by hyphens (e.g., 1234-12-A02-01-01)
        - Identifier 1: '0000' or number > 2000
        - Identifier 2: ## or ##.# or 6## or ###
        - Identifier 3: ### or Letter+numbers or 2Letters+numbers
        - Identifier 3 might be in group 4 or 5
        
        Args:
            part_number: String containing the part number
            
        Returns:
            Dictionary with parsed components or error
        """
        # Grundlæggende validering
        if not isinstance(part_number, str) or not part_number:
            return {'error': 'Invalid input'}
            
        parts = part_number.strip().split('-')
        if len(parts) < 3:
            return {'error': 'Invalid part number format'}
        
        result = {
            'identifier1': '',
            'identifier2': '',
            'identifier3': '',
            'original': part_number
        }
        
        # Handle Identifier 1 (Project Number)
        identifier1 = parts[0]
        if not (identifier1 == '0000' or (identifier1.isdigit() and int(identifier1) > 2000)):
            return {'error': 'Invalid part number'}
        
        result['identifier1'] = identifier1
        
        # Handle Identifier 2 (Area Number or Special Codes)
        identifier2 = parts[1]
        if (len(identifier2) == 2 and identifier2.isdigit()) or \
           (len(identifier2) == 4 and identifier2[2] == '.' and identifier2[0:2].isdigit() and identifier2[3].isdigit()) or \
           (len(identifier2) == 3 and identifier2.isdigit()):
            result['identifier2'] = identifier2
        else:
            return {'error': 'Invalid Identifier 2 format'}
        
        # Handle Identifier 3 (Can be in group 3, 4, or 5)
        def is_valid_identifier3(value):
            if value.isdigit() and len(value) == 3:
                return True
            if len(value) >= 2:
                if value[0].isalpha() and value[1:].isdigit():
                    return True
                if len(value) >= 3 and value[0:2].isalpha() and value[2:].isdigit():
                    return True
            return False
        
        # Check groups 3, 4, and 5 for Identifier 3
        potential_id3 = None
        for i in range(2, min(5, len(parts))):
            if is_valid_identifier3(parts[i].split('.')[0]):  # Handle potential revision numbers
                potential_id3 = parts[i].split('.')[0]
                break
        
        if potential_id3:
            result['identifier3'] = potential_id3
        
        # Extract piping type if present
        if result['identifier3']:
            match = re.match(r'([A-Z]{1,2})\d+', result['identifier3'])
            if match:
                result['piping_type'] = match.group(1)
        
        return result

    def categorize(self, part_number: str) -> Tuple[str, str]:
        """
        Kategoriserer et part number baseret på dets mønster
        Returns: (category, type)
        """
        parsed = self.parse_part_number(part_number)
        logging.info(f"Parsing part number: {part_number}")
        logging.info(f"Parsed components: {parsed}")
        
        # Hvis parsing fejlede, returner Other Parts
        if 'error' in parsed:
            logging.warning(f"Parsing failed for {part_number}: {parsed['error']}")
            return ('Other Parts', 'Other Parts')

        # Check Piping FØRST - højeste prioritet
        if parsed.get('piping_type'):
            piping_type = parsed['piping_type']
            valid_piping_codes = self.config['piping']['identifiers']['identifier3']
            logging.info(f"Checking piping type: {piping_type} against valid codes: {valid_piping_codes}")
            
            # Tjek om det er en Basic Component først
            if parsed['identifier1'] == '0000' and parsed['identifier2'].startswith('3'):
                valid_basic_codes = [str(code) for code in self.config['basic_components']['identifiers']['identifier2']]
                if parsed['identifier2'][:3] in valid_basic_codes:
                    result = (
                        self.config['basic_components']['category'],
                        self.config['basic_components']['types'].get(parsed['identifier2'][:3], 'Unknown')
                    )
                    logging.info(f"Identified as basic component despite piping type: {result}")
                    return result
            
            if piping_type in valid_piping_codes:
                result = (
                    self.config['piping']['category'],
                    self.config['piping']['types'].get(piping_type, 'Unknown')
                )
                logging.info(f"Identified as piping: {result}")
                return result

        # Check Project Specific (PS) - før Basic Components
        if parsed['identifier3'] and parsed['identifier3'].startswith('PS'):
            result = (
                'Project Specific',
                'Project Specific'
            )
            logging.info(f"Identified as project specific: {result}")
            return result

        # Check Basic Components og Suppliers Parts (0000-xxx) - kun identifier1 og identifier2
        if parsed['identifier1'] == '0000':
            area_code = parsed['identifier2'][:3]
            logging.info(f"Checking 0000-xxx number: {area_code}")
            
            # Check 0000-3xx nummer
            if area_code.startswith('3'):
                valid_basic_codes = [str(code) for code in self.config['basic_components']['identifiers']['identifier2']]
                logging.info(f"Checking basic components codes: {valid_basic_codes}")
                if area_code in valid_basic_codes:
                    result = (
                        self.config['basic_components']['category'],
                        self.config['basic_components']['types'].get(area_code, 'Unknown')
                    )
                    logging.info(f"Identified as basic component: {result}")
                    return result
                    
            # Check 0000-7xx nummer
            if area_code.startswith('7'):
                valid_supplier_codes = [str(code) for code in self.config['suppliers_parts']['identifiers']['identifier2']]
                logging.info(f"Checking supplier codes: {valid_supplier_codes}")
                if area_code in valid_supplier_codes:
                    result = (
                        self.config['suppliers_parts']['category'],
                        self.config['suppliers_parts']['types'].get(area_code, 'Unknown')
                    )
                    logging.info(f"Identified as supplier part: {result}")
                    return result
        
        # Check Equipment/Tank drawings (6xx) - alle 3 identifiers
        if parsed['identifier2']:
            area_code = parsed['identifier2'][:3]
            logging.info(f"Checking equipment code: {area_code}")
            if area_code in [str(code) for code in self.config['equipment']['identifiers']['identifier2']]:
                # Tjek om identifier3 matcher et gyldigt præfiks
                for prefix in self.config['equipment']['identifiers']['identifier3']:
                    if parsed['identifier3'].startswith(prefix):
                        result = (
                            self.config['equipment']['categories'].get(area_code, 'Equipment Drawings'),
                            self.config['equipment']['types'].get(area_code, 'Unknown')
                        )
                        logging.info(f"Identified as equipment: {result}")
                        return result
        
        # Check Area Drawings TIL SIDST - alle 3 identifiers
        if parsed['identifier3']:
            matches = [id3 for id3 in self.config['area_drawings']['identifiers']['identifier3'] 
                      if parsed['identifier3'].startswith(id3)]
            logging.info(f"Checking area drawing matches: {matches}")
            if matches:
                best_match = max(matches, key=len)
                result = (
                    self.config['area_drawings']['category'],
                    self.config['area_drawings']['types'].get(best_match, 'Unknown')
                )
                logging.info(f"Identified as area drawing: {result}")
                return result
        
        logging.warning(f"No category match found for {part_number}, defaulting to Other Parts")
        return ('Other Parts', 'Other Parts')

# def test_categorization():
#     """Test function to verify categorization"""
#     parser = CategoryParser()
    
#     test_cases = [
#         "4003-99-BM015-01",
#         "0000-710-001",
#         "4003-99-PS02-01",
#         "0000-301-010",
#         "4003-99-PS08-01",
#         "4003-99-CD001-01",
#         "4003-05-A01",
#         "4003-05.2-A01",
#         "4003-615-A01",
#         "4003-615-BM005",
#         "4003-613-AA001-01",
#         "4003-613-D01",
#         "4003-613-A01",
#         "4003-613-PS01",
#         "4003-630-01-CS001",
#         "4003-615-F01",
#         "asdf-grhf-slkdfj"
#     ]
    
#     print("\nPart Number Categorization Test:")
#     print("-" * 80)
#     print(f"{'Part Number':<20} {'Category':<20} {'Type':<30}")
#     print("-" * 80)
    
#     for part_number in test_cases:
#         category, type_ = parser.categorize(part_number)
#         print(f"{part_number:<20} {category:<20} {type_:<30}")

# if __name__ == "__main__":
#     test_categorization()

