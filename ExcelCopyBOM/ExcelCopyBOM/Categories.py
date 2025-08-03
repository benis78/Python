import csv
import re
import os

class PartNumberParser:
    def __init__(self, csv_file=None):
        if csv_file is None:
            csv_file = os.path.join(os.path.dirname(__file__), "Categories.csv")
        self.categories = self._load_categories(csv_file)
    
    def _load_categories(self, csv_file):
        categories = []
        with open(csv_file, 'r', encoding='utf-8') as file:
            reader = csv.reader(file, delimiter=';')
            for row in reader:
                if len(row) >= 4:  # Sikrer at vi har mindst gruppe1, gruppe2, gruppe3 og kategori
                    categories.append({
                        'gruppe1': row[0].strip(),
                        'gruppe2': row[1].strip(),
                        'gruppe3': row[2].strip(),
                        'kategori': row[3].strip(),
                        'type': row[4].strip() if len(row) > 4 else ''
                    })
        return categories
    
    def parse_part_number(self, part_number):
        # Fjern eventuelle mellemrum og split på bindestreg
        groups = part_number.strip().split('-')
        
        # Initialiser grupperne
        gruppe1 = groups[0] if len(groups) > 0 else ''
        gruppe2 = groups[1] if len(groups) > 1 else ''
        gruppe3 = groups[2] if len(groups) > 2 else ''
        
        return {
            'gruppe1': gruppe1,
            'gruppe2': gruppe2,
            'gruppe3': gruppe3,
            'original': part_number
        }
    
    def _match_pattern(self, value, pattern):
        # Håndterer forskellige mønstre:
        # "2000 - 9999" -> tjekker om tallet er i intervallet
        # "****" -> matcher alt
        # Specifikke mønstre som "A¤¤" -> A efterfulgt af 2 tal
        
        if pattern == '****':
            return True
            
        if ' - ' in pattern:
            # Håndter interval (f.eks. "2000 - 9999")
            start, end = pattern.split(' - ')
            try:
                num_value = int(value)
                return int(start) <= num_value <= int(end)
            except ValueError:
                return False
                
        # Konverter mønster til regex
        pattern = pattern.replace('¤', r'\d')
        return bool(re.match(f'^{pattern}$', value))
    
    def find_category(self, part_number):
        parsed = self.parse_part_number(part_number)
        
        # Speciel håndtering af 0000-3 part numbers
        if parsed['gruppe1'] == '0000' and parsed['gruppe2'].startswith('3'):
            return "Basic Components"
        
        # Først tjek for specifikke gruppe2 matcher (f.eks. 615 for Primary Digester)
        for cat in self.categories:
            if cat['gruppe2'] and cat['gruppe2'] != '****' and self._match_pattern(parsed['gruppe2'], cat['gruppe2']):
                return cat['kategori']
        
        # Derefter tjek for gruppe3 mønstre
        for cat in self.categories:
            if (cat['gruppe3'] and 
                self._match_pattern(parsed['gruppe1'], cat['gruppe1']) and
                (cat['gruppe2'] == '****' or not cat['gruppe2'] or self._match_pattern(parsed['gruppe2'], cat['gruppe2'])) and
                self._match_pattern(parsed['gruppe3'], cat['gruppe3'])):
                return cat['kategori']
        
        return "Other Items"

def test_part_numbers():
    parser = PartNumberParser()
    test_file = os.path.join(os.path.dirname(__file__), "testKategori.csv")
    
    print("Test af Part Numbers:")
    print("-" * 50)
    
    with open(test_file, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for row in reader:
            if row and row[0]:  # Check if row exists and has content
                part_number = row[0].strip()
                category = parser.find_category(part_number)
                print(f"{part_number}: {category}")

if __name__ == "__main__":
    test_part_numbers()
