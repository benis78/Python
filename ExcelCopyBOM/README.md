# ExcelCopyBOM

Et program til at behandle BOM (Bill of Materials) Excel filer, kategorisere komponenter og kopiere tilhørende tegninger.

## Installation

1. Klon dette repository:
```bash
git clone https://github.com/yourusername/ExcelCopyBOM.git
cd ExcelCopyBOM
```

2. Opret et virtuelt miljø (anbefalet):
```bash
python -m venv venv
source venv/bin/activate  # På Linux/Mac
venv\Scripts\activate     # På Windows
```

3. Installer de nødvendige pakker:
```bash
pip install -r requirements.txt
```

## Brug

1. Start programmet:
```bash
python main.py
```

2. I brugergrænsefladen:
   - Vælg BOM Excel fil
   - (Valgfrit) Vælg en tidligere BOM fil til sammenligning
   - Vælg ønskede indstillinger:
     - Inkluder udstyr
     - Inkluder datablad
     - Find revisioner før en bestemt dato
   - Klik på "Start behandling"

3. Programmet vil:
   - Oprette en kopi af BOM filen
   - Kategorisere alle komponenter
   - Oprette separate ark for hver kategori
   - Oprette en samlet stykliste
   - Kopiere relevante tegninger til kategori-mapper

## Filstruktur

```
ExcelCopyBOM/
├── main.py              # Hovedprogram
├── requirements.txt     # Påkrævede pakker
├── Categories.csv       # Kategoridefinitioner
├── steps/              # Programmoduler
│   ├── step1_gui.py    # Brugergrænseflade
│   ├── step2_data.py   # Data indlæsning
│   ├── step3_categorize.py  # Kategorisering
│   ├── step4_partlist.py    # Stykliste
│   ├── step5_format.py      # Formatering
│   └── step6_drawings.py    # Tegningshåndtering
└── logs/               # Logfiler
```

## Kategorier

Programmet kategoriserer komponenter baseret på deres part numbers i følgende grupper:

1. Basic Components (0000-3xx)
2. Piping kategorier (BM, AF, AA, etc.)
3. Area Drawings
4. Tank Drawings
5. Andre kategorier

Se `Categories.csv` for den komplette liste af kategorier og deres definitioner.

## Krav

- Windows 10 eller nyere
- Python 3.8 eller nyere
- Microsoft Excel
- Adgang til tegningsarkiv (lokalt eller netværk)

## Fejlfinding

1. Tjek logfiler i `logs` mappen for detaljerede fejlbeskrivelser
2. Sikr at Excel er installeret og fungerer korrekt
3. Verificer at kildefilen har det korrekte format med påkrævede kolonner

## Support

Ved problemer eller spørgsmål:
1. Tjek logfilen i `logs` mappen
2. Kontakt support på [email/telefon]
3. Opret et issue på GitHub 