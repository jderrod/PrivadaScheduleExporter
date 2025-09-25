@echo off
echo Updating Schedule Extractor DLL...
echo.
echo Please close Revit first, then press any key to continue...
pause

echo Copying updated DLL...
copy "bin\Release\net48\ScheduleExtractor.dll" "C:\Users\james.derrod\AppData\Roaming\Autodesk\Revit\Addins\2026\ScheduleExtractor\"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo DLL updated successfully!
    echo You can now restart Revit and test the add-in.
) else (
    echo.
    echo Failed to update DLL. Make sure Revit is closed.
)

echo.
pause
