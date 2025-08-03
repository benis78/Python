"""
Test af PartNumberCategorizer med nye part numbers
"""

import logging
from steps.step3_categorize import PartNumberCategorizer

def test_new_numbers():
    """Test af kategorisering af nye part numbers"""
    # Konfigurer logging
    logging.basicConfig(level=logging.DEBUG)
    
    # Opret kategoriserer
    categorizer = PartNumberCategorizer()
    
    # Test cases
    test_cases = [
        "4003-99-BM015-01",
        "0000-710-001",
        "4003-99-PS02-01",
        "0000-710-010",
        "4003-99-PS08-01",
        "4003-99-PA001-01",
        "0000-710-001",
        "4003-99-CD001-01",
        "4003-99-CD001-01-06",
        "0000-717-201",
        "4003-05-A01",
        "4003-05.2-A01",
        "4003-615-A01",
        "4003-615-BM005",
        "4003-613-E01",
        "4003-613-D01",
        "4003-613-A01",
        "4003-613-PS01",
        "4003-613-B01",
        "4003-615-F01",
        "4256-613-BF01-01"
    ]
    
    # Header
    print("\nTest resultater:")
    print("-" * 100)
    print("\nPart Number".ljust(40) + "| Type".ljust(30) + "| Kategori".ljust(30))
    print("-" * 100)
    
    # Kør tests
    for part_number in test_cases:
        actual_type, actual_category = categorizer.categorize(part_number)
        print(f"{part_number.ljust(40)}| {actual_type.ljust(30)}| {actual_category}")
        
    print("\n" + "-" * 100)

if __name__ == "__main__":
    test_new_numbers() 