# Multi-filter til Excel

Filtrér en kolonne på mange værdier på én gang. Du indsætter listen i en boks i stedet for at klikke flueben af én for én.

Ingen VBA og ingen internet. Det er én .xll-fil, som du installerer via Excels almindelige Tilføjelsesprogrammer-dialog.

## Download
- **[MultiFilter-64bit.xll](https://github.com/Zeraxdk/multifilter/raw/main/download/MultiFilter-64bit.xll)**: til langt de fleste
- [MultiFilter-32bit.xll](https://github.com/Zeraxdk/multifilter/raw/main/download/MultiFilter-32bit.xll): kun til 32-bit Excel

Er du i tvivl om, hvilken du har? Gå til Filer → Konto → Om Excel. Øverst står der "64-bit" eller "32-bit".

## Installation (uden IT og uden admin)
1. Gem filen et sted, hvor den kan blive liggende, fx `Dokumenter\Excel-tilføjelser\`.
2. **Fjern blokering:** højreklik på filen → Egenskaber → sæt flueben i **Fjern blokering** (Unblock) → OK.
   Windows blokerer .xll-filer, der kommer fra internettet eller en mail.
3. Gå i Excel til Filer → Indstillinger → Tilføjelsesprogrammer. Vælg *Administrer: Excel-tilføjelsesprogrammer* nederst → **Udfør...**.
   Har du fanen Udvikler, kan du også gå til Udvikler → Excel-tilføjelsesprogrammer.
4. Klik **Gennemse...**, vælg `.xll`-filen, og tjek, at **Multi-filter** har flueben → OK.

Nu ligger **Multi-filter** på fanen **Data**, og den bliver indlæst hver gang du starter Excel.
Du fjerner den igen ved at fjerne fluebenet i samme dialog.

## Brug
1. Klik på en celle i den kolonne, du vil filtrere, fx *Material*.
2. Klik på **Multi-filter** på fanen Data.
3. Indsæt dine værdier med **én per linje**. Det er sådan de kommer, når du kopierer en kolonne i Excel.
   Semikolon og tab virker også. Komma bruges **ikke**, fordi det er decimaltegn.
4. Klik **Filtrér** eller tryk **Ctrl+Enter**.

- **Indeholder (delvis match):** `ALU FOIL` finder alle rækker, hvor teksten indgår.
- Hvis nogle værdier ikke findes, bliver vinduet stående med en liste over dem.
- **Ryd kolonne** og **Ryd alle** ligger både på båndet og i vinduet.
- Virker på almindelige filtre og på formaterede tabeller. Hvis der ikke er et filter, bliver det oprettet.
- Excel skal kunne køre .NET Framework 4.7.2. Det er indbygget i Windows 10 og 11.

## Til udvikling
Kildekoden ligger i `xll/` (C# med [Excel-DNA](https://excel-dna.net)). Den bygges med `dotnet build -c Release`.
Det færdige resultat havner i `bin/Release/net472/publish/`. Byg fra en kort mappesti, ellers fejler pakketrinnet.

Den gamle web-version (`taskpane.html`, `manifest.xml`, `Installer.cmd`) ligger her stadig, men den bliver ikke længere vedligeholdt.
