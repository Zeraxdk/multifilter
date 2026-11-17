# Multi-filter: tilføjelsesprogram til Excel (uden VBA)

Et Office-tilføjelsesprogram (samme slags som dem i butikken). Det lægger knappen **Multi-filter** på fanen **Data**.
Knappen åbner et sidepanel, hvor du indsætter værdier og klikker **Filtrér**.

- Separatorer: komma, semikolon, linjeskift, tab eller `|`. Du kan også indsætte en hel kolonne direkte.
- Den kolonne du filtrerer, er den der har den aktive celle. Panelet viser hvilken kolonne det er.
- **Indeholder**: matcher delvist, så fx `114` finder alt der indeholder 114.
- Hvis nogle værdier ikke findes, står de på en liste.
- Virker på almindelige filtre og på formaterede tabeller. Hvis der ikke er et filter, bliver det oprettet.
- Kræver Microsoft 365 eller Excel 2021 eller nyere.

## Filer
| Fil | Hvad |
|---|---|
| `taskpane.html` | Selve panelet (HTML/JS) |
| `assets/icon-*.png` | Ikoner |
| `manifest.xml` | Fortæller Excel hvor panelet ligger og hvor knappen skal sidde |

## 1. Hosting
Panelet ligger på GitHub Pages: https://zeraxdk.github.io/multifilter/taskpane.html
Når du ændrer noget og pusher til `main`, er det opdateret i Excel næste gang panelet åbnes.

(Der ligger ingen data i filerne. Dine Excel-data forlader aldrig din computer.)

## 2. Installér i Excel
**Arbejds-PC (anbefalet):** Bed IT om at udrulle `manifest.xml` via *Microsoft 365 admin center → Integrerede apps*.
Så dukker det op hos dig og eventuelle kolleger, ligesom andre tilføjelsesprogrammer.

**Selv, via en delt mappe:**
1. Læg `manifest.xml` i en mappe, og del den (højreklik → Egenskaber → Deling). Notér stien, fx `\\MIN-PC\addins`.
2. I Excel: Filer → Indstillinger → Center for sikkerhed → Indstillinger for Center for sikkerhed → *Kataloger over pålidelige tilføjelsesprogrammer*.
   Indsæt stien, klik *Tilføj katalog*, sæt flueben i *Vis i menu* og genstart Excel.
3. Gå til Hjem → Tilføjelsesprogrammer → Flere tilføjelsesprogrammer → **DELT MAPPE** → Multi-filter → Tilføj.

**Excel på nettet:** Hjem → Tilføjelsesprogrammer → Flere tilføjelsesprogrammer → Mine tilføjelsesprogrammer → *Upload mit tilføjelsesprogram* → vælg `manifest.xml`.
