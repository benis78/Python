# ExcelCopyBOM Implementation Checklist

## GUI Implementation
- [x] GUI window positioning (topmost, can be overlapped)
- [x] Browse field for "Open Excel BOM List"
- [x] Browse field for "Previous BOM List"
- [x] Checkbox for "Update Index File"
- [x] Start button at bottom
- [x] Progress bar with English descriptions
- [x] "Done" dialog with auto-open of output folder

## Main Processing (BOM Sheet)
- [x] Load Excel BOM file with pandas
- [x] Extract Part Number and REV from filename (REV is last letter before " - ")
- [x] Preserve source file unchanged
- [x] Identify column numbers from header row (Item, Part Number, REV, BOM Structure, Description, QTY, D, t, L)
- [x] Insert arrangement row in row 2:
  - [x] Item = 0
  - [x] Part Number = from filename
  - [x] REV = from filename
  - [x] BOM Structure = None
  - [x] Description = "Arrangement Drawing"
  - [x] QTY = 1
  - [x] D = 1
  - [x] t = 1
  - [x] L = 1
- [x] Delete supplier parts (0000-700, 0000-701, 0000-702)
- [x] Move revision letters to REV column using extract_revision_from_partnumber function
- [x] Group rows by parent/child hierarchy (Multi-Level Numbering System)
- [x] Collapse all rows after grouping
- [x] Handle BOM Structure rules:
  - [x] Delete children if parent is "Inseparable" or starts with "0000-3"
  - [x] Delete row if "Phantom"
- [x] Calculate Total QTY (multiply QTY with parent QTY)
- [x] Copy drawings to category folders
- [x] Update REV column from latest drawings
- [x] Add Drawing status column (DWG~PDF, DWG!PDF, PDF, DWG, or empty)

## Image Handling
- [x] Extract images from "Thumbnail" column
- [x] Handle NULL or empty thumbnail cells
- [x] Save images to temp folder with unique names
- [x] Link images to Part Numbers
- [x] Reinsert images in same column and row (Image height 2.38cm width 2.38cm. Orginal image size: 4,76cm x 4,76cm scale 50%)
- [x] Clean up temp files after processing
- [x] Maintain image formatting (column width and row height)

## Partlist Sheet
- [x] Copy header row from BOM sheet (except Item column)
- [x] Aggregate identical Part Number/REV rows
- [x] Sum Total QTY for identical items
- [x] Copy all rows (unique Part Number/REV combinations)

## Category Sheets
- [x] Create sheets based on categories
- [x] Special handling for Piping category (use Type value instead)
- [x] Copy header row to each sheet
- [x] Copy children rows for Piping based on Type value
- [x] Sort sheets (BOM, Partlist, alphabetical)

## Compare Functionality
- [x] Create Compare sheet with naming format: "Compare OLDREV-NEWREV"
- [x] Copy header row from BOM (except Item column)
- [x] Compare Partlist sheets (NEW vs OLD)
- [x] Handle revision changes:
  - [x] Copy NEW rows with higher REV
  - [x] Copy OLD rows missing in NEW
  - [x] Show changed cell values with old value in parentheses
- [x] Apply color coding:
  - [x] Light green + bold for new rows
  - [x] Light yellow + bold cell for changed values
  - [x] Light red for removed rows
- [x] Copy only changed/new files to category folders
- [x] Sort sheets (BOM, Partlist, Compare, alphabetical)

## Formatting
- [x] Apply row heights:
  - [x] 26px for header row
  - [x] 91px for data rows
- [x] Bold header row
- [x] Add filters to header row
- [x] Freeze top row
- [x] Apply column widths:
  - [x] Item (A): 52px
  - [x] Part Number (B): 152px
  - [x] Rev (C): 47px
  - [x] Description 1 (D): 111px
  - [x] Thumbnail (E): 93px
  - [x] BOM Structure (F): 115px
  - [x] Description (G): 423px
  - [x] Material (H): 135px
  - [x] Title (I): 173px
  - [x] QTY (J): 48px
  - [x] Total QTY (K): 82px
  - [x] Diameter (L): 39px
  - [x] Thickness (M): 39px
  - [x] Length (N): 39px
  - [x] Keywords (O): 200px
  - [x] Type (P): 180px
  - [x] Category (Q): 125px
  - [x] Drawings (R): 94px
- [x] Special formatting for Compare sheet

## Database Integration
- [x] Connect to network database
- [x] Implement file search functionality
- [x] Handle database update option (file_indexer.exe)
- [x] Optimize search queries with caching
- [x] Implement drawing status detection
- [x] Implementer DrawingDatabase klasse
- [x] Tilføj metode til at søge efter tegninger
- [x] Håndter database forbindelser korrekt
- [x] Implementer fejlhåndtering 