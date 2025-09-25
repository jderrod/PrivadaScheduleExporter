@echo off
echo Building Schedule Extractor Add-in...
echo.

REM Build the project
dotnet build --configuration Release
if %ERRORLEVEL% NEQ 0 (
    echo Build failed!
    pause
    exit /b 1
)

echo Build successful!
echo.

REM Create the Revit add-ins directory if it doesn't exist
set "REVIT_ADDINS_DIR=%APPDATA%\Autodesk\Revit\Addins\2026"
set "ADDIN_DIR=%REVIT_ADDINS_DIR%\ScheduleExtractor"

if not exist "%REVIT_ADDINS_DIR%" (
    echo Creating Revit add-ins directory...
    mkdir "%REVIT_ADDINS_DIR%"
)

if not exist "%ADDIN_DIR%" (
    echo Creating add-in directory...
    mkdir "%ADDIN_DIR%"
)

REM Copy the manifest file
echo Copying add-in manifest...
copy "ScheduleExtractor.addin" "%REVIT_ADDINS_DIR%\"
if %ERRORLEVEL% NEQ 0 (
    echo Failed to copy manifest file!
    pause
    exit /b 1
)

REM Copy the DLL files
echo Copying add-in assemblies...
copy "bin\Release\net48\ScheduleExtractor.dll" "%ADDIN_DIR%\"
copy "bin\Release\net48\Newtonsoft.Json.dll" "%ADDIN_DIR%\"

if %ERRORLEVEL% NEQ 0 (
    echo Failed to copy assembly files!
    pause
    exit /b 1
)

echo.
echo Installation complete!
echo.
echo The add-in has been installed to:
echo   %REVIT_ADDINS_DIR%
echo.
echo You can now:
echo 1. Start Revit 2026
echo 2. Open your project file
echo 3. Go to the Add-ins tab
echo 4. Click "Export Schedules to JSON"
echo.
echo The exported JSON files will be saved in a "ScheduleExports" folder
echo in the same directory as your Revit project.
echo.
pause
