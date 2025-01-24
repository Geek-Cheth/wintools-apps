# Script to create executable from RAMOptimizer.ps1
Write-Host "Creating RAM Optimizer Executable..." -ForegroundColor Cyan

# Check if PS2EXE module is installed
if (-not (Get-Module -ListAvailable -Name PS2EXE)) {
    Write-Host "Installing PS2EXE module..." -ForegroundColor Yellow
    Install-Module -Name PS2EXE -Force -Scope CurrentUser
}

# Import the module
Import-Module PS2EXE

# Create the executable
try {
    Write-Host "Converting PowerShell script to executable..." -ForegroundColor Yellow
    $params = @{
        InputFile = "RAMOptimizer.ps1"
        OutputFile = "RAMOptimizer.exe"
        NoConsole = $false
        RequireAdmin = $true
    }
    
    Invoke-PS2EXE @params

    Write-Host "Executable created successfully!" -ForegroundColor Green
    Write-Host "Output file: $(Resolve-Path RAMOptimizer.exe)" -ForegroundColor Cyan
}
catch {
    Write-Host "Error creating executable: $($_.Exception.Message)" -ForegroundColor Red
} 