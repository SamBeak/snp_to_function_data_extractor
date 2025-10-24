@echo off
chcp 65001 >nul
echo ========================================
echo Gene Automation Setup
echo ========================================
echo.

echo Checking and installing required libraries...
echo.

pip install requests beautifulsoup4 openpyxl

echo.
echo ========================================
echo Starting gene automation...
echo ========================================
echo.

python gene_automation.py

echo.
echo ========================================
echo Process completed!
echo ========================================
echo.

if exist "gene_data_output.xlsx" (
    echo Output file created: gene_data_output.xlsx
    echo.
    set /p OPEN="Do you want to open the Excel file? (Y/N): "
    if /i "%OPEN%"=="Y" start "" "gene_data_output.xlsx"
)

echo.
pause
