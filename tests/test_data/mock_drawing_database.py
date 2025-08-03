class MockDrawingDatabase:
    """Mock version af DrawingDatabase til test"""
    def __init__(self):
        self.drawings = {
            '1234-01': [
                {'filepath': 'drawings/1234-01A.dwg', 'filename': '1234-01A.dwg'},
                {'filepath': 'drawings/1234-01A.pdf', 'filename': '1234-01A.pdf'}
            ],
            '1234-02': [
                {'filepath': 'drawings/1234-02B.dwg', 'filename': '1234-02B.dwg'}
            ]
        }
        
    def find_drawings(self, part_number: str) -> list:
        """Find tegninger for et part number"""
        return self.drawings.get(part_number, []) 