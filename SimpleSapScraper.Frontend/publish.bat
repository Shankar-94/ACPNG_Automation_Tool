@echo off
echo Publishing ACPNG Automation Tool...
echo.

cd /d "%~dp0SimpleSapScraper.Frontend"

dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true -p:EnableCompressionInSingleFile=false

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Publish failed!
    pause
    exit /b 1
)

echo.
echo Publish complete!
echo.
echo Your .exe is here:
echo SimpleSapScraper.Frontend\bin\Release\net10.0-windows\win-x64\publish\SimpleSapScraper.Frontend.exe
echo.

start "" "SimpleSapScraper.Frontend\bin\Release\net10.0-windows\win-x64\publish"

pause