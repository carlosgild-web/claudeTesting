# HPI QA Agent - local server launcher
# Run from the QA_Agent directory: .\start.ps1

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

Write-Host ""
Write-Host "  HPI QA Agent  -  localhost:5050" -ForegroundColor Cyan
Write-Host ""

# Check ANTHROPIC_API_KEY
if (-not $env:ANTHROPIC_API_KEY) {
    Write-Host "  ERROR: ANTHROPIC_API_KEY is not set." -ForegroundColor Red
    Write-Host '  Run: $env:ANTHROPIC_API_KEY = "sk-ant-..."' -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# Install dependencies if needed
$missing = $false
foreach ($pkg in @("flask", "flask_cors", "anthropic", "pandas", "openpyxl")) {
    python -c "import $pkg" 2>$null
    if ($LASTEXITCODE -ne 0) { $missing = $true; break }
}

if ($missing) {
    Write-Host "  Installing dependencies..." -ForegroundColor Yellow
    pip install -r requirements.txt
    Write-Host ""
}

Write-Host "  Starting server - open a dashboard and click Ask the Data" -ForegroundColor Green
Write-Host "  Press Ctrl+C to stop." -ForegroundColor Gray
Write-Host ""

python server.py
