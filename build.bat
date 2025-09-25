@echo off
echo Building Schedule Extractor Add-in...
echo.

dotnet build --configuration Release

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
    echo Output files are in: bin\Release\net48\
) else (
    echo.
    echo Build failed!
)

echo.
pause
