from Categories import CategoryParser
import unittest

class TestPartNumberParsing(unittest.TestCase):
    def setUp(self):
        self.parser = CategoryParser()

    def test_part_number_parsing(self):
        test_cases = [
            ("4003-99-BM015-01", {    # Project with piping
                'identifier1': '4003',
                'identifier2': '99',
                'identifier3': 'BM015',
                'piping_type': 'BM',
                'original': '4003-99-BM015-01'
            }),
            ("0000-710-001", {        # Suppliers part
                'identifier1': '0000',
                'identifier2': '710',
                'identifier3': '001',
                'original': '0000-710-001'
            }),
            ("4003-05.2-A01", {       # Project with sub-area
                'identifier1': '4003',
                'identifier2': '05.2',
                'identifier3': 'A01',
                'original': '4003-05.2-A01'
            }),
            ("4003-615-A01", {        # Equipment/Tank drawing
                'identifier1': '4003',
                'identifier2': '615',
                'identifier3': 'A01',
                'original': '4003-615-A01'
            })
        ]
        
        for part_number, expected in test_cases:
            with self.subTest(part_number=part_number):
                result = self.parser.parse_part_number(part_number)
                self.assertEqual(result, expected, 
                    f"\nPart number: {part_number}\nExpected: {expected}\nGot: {result}")

if __name__ == '__main__':
    unittest.main(verbosity=2)