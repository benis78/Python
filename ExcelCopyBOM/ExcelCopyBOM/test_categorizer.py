"""
Test af PartNumberCategorizer
"""

from steps.step3_categorize import PartNumberCategorizer
import logging

# Konfigurer logging
logging.basicConfig(level=logging.DEBUG)

def test_categorization():
    """Test kategorisering af part numbers"""
    categorizer = PartNumberCategorizer()
    
    # Test cases
    test_cases = [
        # Basic Components og Suppliers Parts
        ("0000-301-001", "Tank Connection", "Basic Components"),
        ("0000-700-002", "Equipment", "Suppliers Parts"),
        
        # Area Drawings
        ("4003-04.1-B01", "Area Building Drawing", "Area Drawings"),
        ("4003-04.1-A01", "Area Arrangement Drawing", "Area Drawings"),
        ("4003-04.1-E01", "Area Equipment Drawing", "Area Drawings"),
        ("4003-04.1-F01", "Area Foundation Drawing", "Area Drawings"),
        ("4003-04.1-PS03-01", "Project Specific Parts", "Area Drawings"),
        
        # Piping
        ("4003-04.1-BM004-01", "Biomass Piping", "Piping"),
        ("4003-04.1-BM004-01-01", "Biomass Piping", "Piping"),
        ("4003-04.1-BM020-02", "Biomass Piping", "Piping"),
        ("4003-04.1-CD005-01", "Cable Ducts", "Piping"),
        ("4003-04.1-HW004-01", "Hot Water Piping", "Piping"),
        ("4003-615-01-AF001-01", "Anti Foam Piping", "Piping"),
        
        # Equipment
        ("4003-615-A01", "Primary Digester", "Tank Drawings"),
        ("4003-621-A04", "Heat-Exchangers", "Equipment Drawings"),
        ("4003-630-02-A02", "GasMix", "Tank Drawings"),
        
        # Ugyldige numre
        ("BC-000170", "Unknown", "Unknown"),
        ("1234-01", "Unknown", "Unknown"),
        ("not-a-number", "Unknown", "Unknown")
    ]
    
    print("\nTest resultater:")
    print("-" * 140)
    print("\nPart Number".ljust(40) + "| Forventet Type".ljust(30) + "| Faktisk Type".ljust(30) + 
          "| Forventet Kategori".ljust(30) + "| Faktisk Kategori".ljust(30) + "| Status")
    print("-" * 140 + "\n")
    
    passed = 0
    for part_number, expected_type, expected_category in test_cases:
        actual_type, actual_category = categorizer.categorize(part_number)
        status = "✓" if actual_type == expected_type and actual_category == expected_category else "✗"
        if status == "✓":
            passed += 1
            
        print(f"{part_number}".ljust(40) + 
              f"| {expected_type}".ljust(30) + 
              f"| {actual_type}".ljust(30) +
              f"| {expected_category}".ljust(30) + 
              f"| {actual_category}".ljust(30) + 
              f"| {status}")
        
    print("\n" + "-" * 140)
    print(f"\nResultat: {passed}/{len(test_cases)} tests bestået ({(passed/len(test_cases))*100:.1f}%)\n")

if __name__ == "__main__":
    test_categorization() 