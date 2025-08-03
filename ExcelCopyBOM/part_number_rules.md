# Part Number Kategorisering Regler

## Grundlæggende Struktur
Et Part Number består af flere grupper adskilt af bindestreger (-).

## Gruppe 1 (Første 4 tegn)
- Altid 4 alfanumeriske tegn
- Mulige værdier:
  - `0000`: Suppliers Parts eller Basic Components
  - `2000-9999`: Projekt numre
  - `****`: Sales (hvis de første 2 eller 4 tegn er bogstaver, f.eks. "DK83")

## Gruppe 2
Består af følgende muligheder:
- 2 eller 4 alfanumeriske tegn = "Area Drawings"
- 3 tal startende med 3 = "Basic Components"
- 3 tal startende med 6 = "Basic Equipment"

## Gruppe 3 (op til 4 alfanumeriske tegn)
Kategorisering baseret på tidligere grupper:

### Hvis Gruppe 1 er "0000":
- Starter med 3 = "Basic Components"
- Starter med 7 = "Suppliers Parts"

### Hvis Gruppe 1+2 er "Basic Equipment":
- Starter med bogstav = "Basic Equipment Drawing"

### Hvis Gruppe 1+2 er "Area Drawings":
Starter med:
- `A` = Arrangement Drawing
- `E` = Equipment Drawing
- `P` eller `S` = Sub Terrain Piping Plan
- `F` = Foundation
- `PS` = Project Specific Parts
- `B` = Building

### Piping Kategorier (hvis Gruppe 1+2 er "Area Drawings"):
- `BM` = Biomass Piping
- `AF` = Anti Foam Piping
- `AA` = Atmospheric Air Piping
- `BG` = Biogas Piping
- `CD` = Cable Ducts
- `CS` = Condensate Piping
- `EL` = Power Cable Piping
- `CO` = Cooling Water Piping
- `EZ` = Enzyme Piping
- `HW` = Hot Water Piping
- `HO` = Hydraulic Oil Piping
- `IC` = Iron Chloride Piping
- `OA` = Odour Piping
- `NT` = Nutrient Piping
- `OG` = Offgas Piping
- `PW` = Potable Water Piping
- `PA` = Pressurized Air Piping
- `RW` = Rain Water Piping
- `SL` = Sulphurous liquid Piping
- `TD` = Technical Drainage Piping
- `TW` = Technical Water Piping

## Efterfølgende Grupper
- Alle efterfølgende grupper er sekventielle løbenumre
- Der kan være flere grupper med løbenumre

## Specialtilfælde
- Alle "Part Number" som ikke passer ind i ovenstående skal i "Other Parts"

## Symbol Forklaring
- `¤` = Tal
- `*` = Alfanumerisk
- `¤¤¤¤` = 4 tal
- `****` = 4 tal, bogstaver eller symboler
- Når der står 2000-9999, betyder det et tal mellem 2000 og 9999
- 100-200 betyder et tal mellem 100 og 200 

## Forklaring af nummersystem
Et Part Number defineres som minimum 3 grupper der er afskilt af bindestreg (F.eks. 1234-12-A02-01-01)
Første gruppe (Identifier 1) er enten 4 nul 0000 eller større end 2000
Anden gruppe (Identifier 2) kan være ## (2 tal f.eks. 02) eller ##.# (02.1) for Area Number. Det kan også starte med tallet 6. Indentifier 2 kan også være ### (3 tal).
Tredje gruppe (Identifier 3) kan være ### (3 tal). Det kan være et bogstav efterfulgt af tal. Det kan være 2 bogstaver efterfulgt af tal. 
der kan være flere grupper og nogle gange kan Identifier 3 ligger i gruppe 4 (efter 3. bindestreg) eller 5 (efter 4. bindestreg) 
Ud fra de første 3 grupper er det muligt at tildele et hvert nummer en Catagories og Types
