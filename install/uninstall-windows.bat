@echo off
chcp 65001 >nul
rem Fjerner PPTimagePoint fra PowerPoint for Windows.
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "9bfd5010-5050-433e-9166-cedb07dc32f8" /f >nul 2>&1
if exist "%LOCALAPPDATA%\PPTimagePoint" rmdir /S /Q "%LOCALAPPDATA%\PPTimagePoint"
echo.
echo PPTimagePoint er fjernet. Start PowerPoint på nytt.
echo.
pause
