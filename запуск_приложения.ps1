# Encoding: UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Obrabotka dannyx 112 i Google Sheets"

# Colors
$SuccessColor = "Green"
$ErrorColor = "Red"
$InfoColor = "Cyan"
$WarningColor = "Yellow"

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Header
Clear-Host
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "       OBRABOTKA DANNYX 112 I GOOGLE SHEETS" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Navigate to script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-ColorOutput "Rabochaya direktoriya: $ScriptDir" $InfoColor
Write-Host ""

# Check virtual environment
if (Test-Path ".venv\Scripts\Activate.ps1") {
    Write-ColorOutput "Aktivaciya virtualnogo okruzheniya..." $InfoColor
    & ".venv\Scripts\Activate.ps1"
    Write-ColorOutput "Virtualnoe okruzhenie aktivirovano" $SuccessColor
} else {
    Write-ColorOutput "Virtualnoe okruzhenie ne naydeno" $WarningColor
    Write-ColorOutput "Prodolzhaem bez aktivacii..." $WarningColor
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-ColorOutput "Zapusk prilozheniya..." $InfoColor
Write-Host ""

# Run application
try {
    python process_data_app.py
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-ColorOutput "Prilozhenie zaversheno uspeshno" $SuccessColor
    } else {
        Write-Host ""
        Write-ColorOutput "Prilozhenie zaversheno s oshibkoy (kod: $LASTEXITCODE)" $ErrorColor
        Write-Host ""
        Write-ColorOutput "Nazhmite lyubuyu klavishu dlya vyxoda..." $WarningColor
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
} catch {
    Write-Host ""
    Write-ColorOutput "OSHIBKA PRI ZAPUSKE:" $ErrorColor
    Write-ColorOutput $_.Exception.Message $ErrorColor
    Write-Host ""
    Write-ColorOutput "Nazhmite lyubuyu klavishu dlya vyxoda..." $WarningColor
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Write-Host ""
