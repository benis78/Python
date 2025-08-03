import pandas as pd

# Opret test data
data = {
    'Item': [1, 2, 3, 4, 5],
    'Part Number': ['1234-01A', '1234-02B', '0000-700-123', '1234-03C', '1234-04D'],
    'REV': ['A', 'B', '', 'C', 'D'],
    'BOM Structure': ['1', '1.1', '2', '2.1', 'Phantom'],
    'Description': ['Part 1', 'Part 2', 'Part 3', 'Part 4', 'Part 5'],
    'QTY': [1, 2, 1, 2, 1],
    'D': [10, 20, 30, 40, 50],
    't': [1, 2, 3, 4, 5],
    'L': [100, 200, 300, 400, 500]
}

# Opret DataFrame
df = pd.DataFrame(data)

# Gem som Excel fil
df.to_excel('test_bom-A -- Test.xlsx', index=False) 