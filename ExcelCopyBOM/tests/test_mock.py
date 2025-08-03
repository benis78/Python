"""
Mock klasser til brug i tests
"""
from unittest.mock import MagicMock

class MockGUI:
    """Mock af ExcelCopyBOMGUI klassen"""
    def __init__(self):
        self.bom_path = MagicMock()
        self.bom_path.get = MagicMock(return_value="test_data.xlsx")
        
        self.prev_bom_path = MagicMock()
        self.prev_bom_path.get = MagicMock(return_value=None)
        
        self.update_index = MagicMock()
        self.update_index.get = MagicMock(return_value=False)
        
        self.progress_var = MagicMock()
        self.status_var = MagicMock()
        self.start_button = MagicMock()
        
    def update_progress(self, value, message):
        """Mock progress update"""
        self.progress_var.set = MagicMock(return_value=value)
        self.status_var.set = MagicMock(return_value=message)
        
    def show_error(self, message):
        """Mock error dialog"""
        pass
        
    def show_completion_dialog(self, output_path):
        """Mock completion dialog"""
        pass 