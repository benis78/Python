# Description Splitter

Dette program splitter beskrivelser i Excel-filer baseret på specifikke regler.

## Funktioner

- Læser Excel-filer (.xlsx, .xls)
- Finder kolonnerne "Description" og "Part Number"
- Splitter beskrivelser i grupper med " - " som separator
- Ekstraherer størrelsesbetegnelser (DN, Ø/ø, SDR, PN, DVR)
- Kopierer data til nye kolonner baseret på regler
- **Sheet-vælger**: Vælg mellem alle sheets eller specifikke sheets
- **Auto-filter**: Automatisk filter-funktionalitet på alle kolonner
- **Formatering**: Professionel formatering af header-rækker

## Regler

1. **Part Number filter**: Kun rækker hvor Part Number starter med "0000-7" behandles
2. **Beskrivelse splitting**: Splitter på " - " (mellemrum bindestreg mellemrum)
3. **Designation**: Første gruppe kopieres til en ny "Designation" kolonne
4. **Størrelsesbetegnelser**:
   - **Ø/ø**: Kopieres til kolonne med header "D" (kun tallet, ikke Ø tegnet)
   - **DN, SDR, PN, DVR**: Får deres egen kolonne med betegnelsen som navn

## Størrelsesbetegnelser der genkendes

- `DN250` → DN kolonne
- `Ø114,3` eller `ø114,3` → D kolonne (kun "114,3")
- `SDR17` → SDR kolonne
- `PN10` → PN kolonne
- `DVR25` → DVR kolonne

## Installation

### Krævede pakker:
```bash
pip install openpyxl
```

### Kør programmet:
```bash
python description_splitter.py
```

### Byg eksekverbar fil:
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --name "DescriptionSplitter" --clean description_splitter.py
```

## Brug

1. Åbn programmet
2. Vælg Excel-fil via "Browse" knappen
3. **Vælg sheets**: Dropdown menu til at vælge "All Sheets" eller specifikke sheets
4. Klik "Start Processing"
5. Programmet gemmer den behandlede fil med "_processed" tilføjet til navnet

## Output

- Opretter nye kolonner efter den sidste eksisterende kolonne
- "Designation" kolonne indeholder første gruppe fra beskrivelsen
- Størrelsesbetegnelser placeres i deres respektive kolonner
- **Formatering**: Header-rækker formateres med fed skrift og grå baggrund
- **Auto-filter**: Alle kolonner får filter-funktionalitet
- Kolonnebredder justeres automatisk
- Kun valgte sheets behandles

## Eksempel

**Input beskrivelse:**
```
Pipe Holder - DN250 - Ø114,3 - PN10
```

**Output:**
- Designation kolonne: "Pipe Holder"
- D kolonne: "114,3" (kun tallet, ikke Ø tegnet)
- DN kolonne: "DN250"  
- PN kolonne: "PN10"

## Moduler der bruges

- `openpyxl` - Excel fil håndtering
- `re` - Regular expressions til størrelsesbetegnelser
- `os` - Operativsystem funktioner
- `pathlib.Path` - Filsti håndtering
- `tkinter` - GUI framework
- `tkinter.filedialog` - Fil vælger dialog
- `tkinter.messagebox` - Besked bokse
- `tkinter.ttk` - Moderne GUI widgets
- `openpyxl.styles` - Excel formatering

## Windows Defender Problemer

Se `README_Windows_Defender.md` for løsninger til Windows Defender problemer med .exe filer. 