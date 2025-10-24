# Gene Automation Script Runner
# This script runs the gene_automation.py script

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Gene Automation Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python and try again." -ForegroundColor Yellow
    pause
    exit 1
}

# Install required libraries
Write-Host ""
Write-Host "Checking and installing required libraries..." -ForegroundColor Yellow
Write-Host ""

try {
    python -m pip install --upgrade pip | Out-Null
    python -m pip install requests beautifulsoup4 openpyxl
    Write-Host "Required libraries installed successfully!" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not install some libraries. Continuing anyway..." -ForegroundColor Yellow
}

# Check if snps.json exists
if (-Not (Test-Path "snps.json")) {
    Write-Host "Error: snps.json file not found!" -ForegroundColor Red
    Write-Host "Please make sure snps.json is in the same directory." -ForegroundColor Yellow
    pause
    exit 1
}

# Check if gene_automation.py exists
if (-Not (Test-Path "gene_automation.py")) {
    Write-Host "Error: gene_automation.py file not found!" -ForegroundColor Red
    pause
    exit 1
}

Write-Host ""
Write-Host "Starting gene automation process..." -ForegroundColor Yellow
Write-Host ""

# Run the Python script
try {
    python gene_automation.py

    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "Process completed successfully!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green

        # Check if output file was created
        if (Test-Path "gene_data_output.xlsx") {
            Write-Host ""
            Write-Host "Output file created: gene_data_output.xlsx" -ForegroundColor Green

            # Ask if user wants to open the file
            $response = Read-Host "Do you want to open the Excel file? (Y/N)"
            if ($response -eq "Y" -or $response -eq "y") {
                Start-Process "gene_data_output.xlsx"
            }
        }
    } else {
        Write-Host ""
        Write-Host "Error: Process failed with exit code $LASTEXITCODE" -ForegroundColor Red
    }
} catch {
    Write-Host ""
    Write-Host "Error occurred while running the script:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
