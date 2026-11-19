@echo off
chcp 65001 >nul
rem Installerer PPTimagePoint i PowerPoint for Windows (gjelder bare din bruker).
set "DEST=%LOCALAPPDATA%\PPTimagePoint"
if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%~dp0manifest.xml" "%DEST%\manifest.xml" >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "9bfd5010-5050-433e-9166-cedb07dc32f8" /t REG_SZ /d "%DEST%\manifest.xml" /f >nul
echo.
echo PPTimagePoint er installert.
echo Lukk PowerPoint helt og start det på nytt.
echo Gå så til Hjem ^> Tillegg ^> Mine tillegg og velg PPTimagePoint.
echo.
pause
