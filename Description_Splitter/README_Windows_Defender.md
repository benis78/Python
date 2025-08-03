# Undgå Windows Defender Problemer

## Metode 1: Unblock filen (Anbefalet)

1. **Højreklik** på `DescriptionSplitter.exe`
2. **Vælg "Properties"**
3. **Tjek "Unblock"** boksen nederst i vinduet
4. **Klik "Apply" og "OK"**

## Metode 2: Tilføj til Windows Defender undtagelser

1. **Åbn Windows Security**
2. **Gå til "Virus & threat protection"**
3. **Klik "Manage settings"**
4. **Under "Exclusions" klik "Add or remove exclusions"**
5. **Tilføj mappen** hvor .exe filen ligger

## Metode 3: Byg med optimale indstillinger

### Brug batch-filen:
```bash
build_exe.bat
```

### Eller manuelt:
```bash
pyinstaller --onefile --noconsole --name "DescriptionSplitter" --clean description_splitter.py
```

## Metode 4: Digitale signering (Avanceret)

Hvis du har et code signing certificate:
```bash
pyinstaller --onefile --noconsole --name "DescriptionSplitter" description_splitter.py
signtool sign /f certificate.pfx /p password DescriptionSplitter.exe
```

## Hvorfor sker dette?

- **Windows Defender** mistænker alle .exe filer der ikke er digitalt signeret
- **PyInstaller** pakker Python-koden ind i en .exe, hvilket kan se mistænkeligt ud
- **Dette er normalt** og ikke et tegn på at filen er farlig

## Sikkerhedstips

1. **Kør kun .exe filer** du har bygget selv eller fra betroede kilder
2. **Scan filen** med antivirus software hvis du er i tvivl
3. **Kør Python-scriptet direkte** hvis du er bekymret: `python description_splitter.py`

## Fejlfinding

Hvis programmet stadig bliver blokeret:
1. **Tjek Windows Event Viewer** for detaljerede fejl
2. **Prøv at køre som administrator**
3. **Temporært deaktiver Windows Defender** (kun for test)
4. **Brug Python-scriptet direkte** i stedet for .exe 