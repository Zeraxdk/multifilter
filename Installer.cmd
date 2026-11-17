@echo off
rem Installerer Multi-filter i Excel for den aktuelle bruger (ingen admin-rettigheder noedvendige).
setlocal
set "DIR=%LOCALAPPDATA%\MultiFilter"
set "MAN=%DIR%\manifest.xml"

echo Installerer Multi-filter...
if not exist "%DIR%" mkdir "%DIR%"

curl.exe -fsSL "https://zeraxdk.github.io/multifilter/manifest.xml" -o "%MAN%"
if errorlevel 1 (
  echo.
  echo Kunne ikke hente manifest.xml. Tjek internetforbindelsen.
  pause
  exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "e77f017a-a6fb-443d-8474-fcf0b1ee13ca" /t REG_SZ /d "%MAN%" /f >nul
if errorlevel 1 (
  echo Kunne ikke registrere tilfoejelsesprogrammet.
  pause
  exit /b 1
)

echo.
echo Faerdig! Luk Excel helt og aabn det igen.
echo Knappen "Multi-filter" ligger paa fanen Data.
pause
