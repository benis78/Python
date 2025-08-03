def group_by_parent_items(sheet):
    """Grupperer rækker baseret på parent items med hierarkisk struktur."""
    last_row = sheet.UsedRange.Rows.Count
    if last_row <= 1:
        return
    
    # Opbyg hierarkisk struktur
    hierarchy = {}  # item_number -> [start_row, end_row, level]
    current_items = {}  # level -> item_number
    
    for row in range(2, last_row + 1):
        item_number = str(sheet.Cells(row, 1).Value)
        if not item_number:
            continue
        
        # Bestem niveau baseret på antal punktummer
        level = len(item_number.split('.'))
        
        # Find parent
        if level > 1:
            parent_number = '.'.join(item_number.split('.')[:-1])
            if parent_number in hierarchy:
                # Opdater parent's end_row
                hierarchy[parent_number][1] = row
        
        # Gem dette item
        hierarchy[item_number] = [row, row, level]
        current_items[level] = item_number
    
    # Opret grupper for hvert niveau, startende med det dybeste
    max_level = max(item[2] for item in hierarchy.values()) if hierarchy else 0
    
    for level in range(max_level, 1, -1):
        for item_number, (start_row, end_row, item_level) in hierarchy.items():
            if item_level == level - 1:  # Parent niveau
                # Find alle direkte children
                children = [child for child in hierarchy.keys() 
                          if child.startswith(item_number + '.') and 
                          len(child.split('.')) == level]
                
                if children:
                    # Find start og slut række for denne gruppe
                    group_start = min(hierarchy[child][0] for child in children)
                    group_end = max(hierarchy[child][1] for child in children)
                    
                    # Opret gruppe
                    range_to_group = sheet.Range(f"{group_start}:{group_end}")
                    range_to_group.Rows.Group()
                    
                    # Opdater parent's end_row hvis nødvendigt
                    hierarchy[item_number][1] = max(hierarchy[item_number][1], group_end)
    
    # Vis alle niveauer som standard
    sheet.Outline.ShowLevels(RowLevels=max_level)

def should_include_excel_row(item_number, sheet):
    """Afgør om en række skal inkluderes baseret på BOM Structure."""
    if not item_number or '.' not in str(item_number):
        return True
    
    # Find den aktuelle række
    current_range = sheet.Range("A:A").Find(item_number)
    if current_range:
        current_structure = str(sheet.Cells(current_range.Row, 5).Value).strip().upper()  # Kolonne E (BOM Structure)
        if current_structure == "INSEPARABLE":
            return True
    
    # Split item number i dele
    parts = item_number.split('.')
    current_parent = ""
    for part in parts[:-1]:
        if current_parent:
            current_parent += "."
        current_parent += part
        
        # Find parent row
        parent_range = sheet.Range("A:A").Find(current_parent)
        if parent_range:
            parent_row = parent_range.Row
            parent_structure = str(sheet.Cells(parent_row, 5).Value).strip().upper()
            if parent_structure == "INSEPARABLE":
                return False
    
    return True

def should_include_row(row, df):
    """Afgør om en række skal inkluderes baseret på BOM Structure (pandas version)."""
    item_number = str(row.iloc[0])
    
    # Find BOM Structure kolonnen
    bom_structure_col = None
    for col in df.columns:
        if "BOM STRUCTURE" in str(col).upper():
            bom_structure_col = col
            break
    
    if bom_structure_col is None:
        return True  # Hvis kolonnen ikke findes, inkluder alle rækker
    
    # Tjek om denne række er markeret som "Inseparable"
    current_structure = str(row[bom_structure_col]).strip().upper()
    if current_structure == "INSEPARABLE":
        return True
    
    # Tjek om denne række er under en "Inseparable" række
    if '.' in item_number:
        parts = item_number.split('.')
        current_parent = ""
        for part in parts[:-1]:
            if current_parent:
                current_parent += "."
            current_parent += part
            
            # Find parent rækken
            parent_mask = df.iloc[:, 0] == current_parent
            if parent_mask.any():
                parent_row = df[parent_mask].iloc[0]
                parent_structure = str(parent_row[bom_structure_col]).strip().upper()
                if parent_structure == "INSEPARABLE":
                    return False
    
    return True

def group_by_parent_items_vba_style(sheet):
    """Grupperer rækker baseret på VBA-implementeringen."""
    try:
        last_row = sheet.UsedRange.Rows.Count
        if last_row <= 1:
            return
        
        # Opbyg group map ligesom i VBA
        group_map = []  # Liste af tupler (cell_address, subgroup_level)
        
        # Find alle rækker og deres gruppe niveauer
        for row in range(2, last_row + 1):
            item_number = str(sheet.Cells(row, 1).Value)
            if not item_number:
                continue
            
            # Beregn subgroup level (antal punktummer)
            subgroup_level = len(item_number.split('.')) - 1
            cell_address = sheet.Cells(row, 1).Address
            group_map.append((cell_address, subgroup_level))
        
        if not group_map:
            return
        
        # Find max subgroup level
        max_subgroup = max(level for _, level in group_map)
        
        # Fjern eksisterende grupper
        try:
            for _ in range(10):  # Samme som VBA koden
                sheet.Range(f"2:{last_row}").Rows.Ungroup()
        except:
            pass  # Ignorer fejl ved ungroup
        
        # Opret grupper fra højeste til laveste niveau
        for current_level in range(max_subgroup, 0, -1):
            start_group = None
            last_group = None
            
            for address, level in group_map:
                if level >= current_level:
                    if start_group is None:
                        start_group = sheet.Range(address)
                    last_group = sheet.Range(address)
                else:
                    if start_group and last_group:
                        # Opret gruppe fra start til slut
                        group_range = sheet.Range(f"{start_group.Row}:{last_group.Row}")
                        group_range.Rows.Group()
                    start_group = None
                    last_group = None
            
            # Håndter sidste gruppe hvis den eksisterer
            if start_group and last_group:
                group_range = sheet.Range(f"{start_group.Row}:{last_group.Row}")
                group_range.Rows.Group()
        
        # Vis alle niveauer
        sheet.Outline.ShowLevels(RowLevels=max_subgroup + 1)
        
    except Exception as e:
        print(f"Fejl i group_by_parent_items_vba_style: {str(e)}")
        raise

def group_by_parent_items(sheet):
    """Wrapper funktion der bruger VBA-style gruppering."""
    return group_by_parent_items_vba_style(sheet) 