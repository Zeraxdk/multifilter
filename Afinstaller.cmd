@echo off
rem Fjerner Multi-filter fra Excel for den aktuelle bruger.
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "e77f017a-a6fb-443d-8474-fcf0b1ee13ca" /f >nul 2>&1
if exist "%LOCALAPPDATA%\MultiFilter" rmdir /s /q "%LOCALAPPDATA%\MultiFilter"
echo Multi-filter er fjernet. Genstart Excel.
pause
