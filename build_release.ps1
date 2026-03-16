Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pythonCandidates = @(
    (Join-Path $projectRoot ".venv-pack\Scripts\python.exe"),
    (Join-Path $projectRoot ".venv311\Scripts\python.exe"),
    "python"
)

$pythonExe = $null
foreach ($candidate in $pythonCandidates) {
    if ($candidate -eq "python") {
        $pythonExe = $candidate
        break
    }
    if (Test-Path $candidate) {
        $pythonExe = $candidate
        break
    }
}

if (-not $pythonExe) {
    throw "No usable Python interpreter was found."
}

Write-Host "[1/3] Building desktop app with PyInstaller..." -ForegroundColor Cyan
& $pythonExe -m PyInstaller --noconfirm --clean QuantumBot.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed with exit code: $LASTEXITCODE"
}

$distRoot = Join-Path $projectRoot "dist\QuantumBot"
$scriptsDist = Join-Path $distRoot "scripts"

Write-Host "[2/3] Creating scripts folder in final dist..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $scriptsDist | Out-Null

$filesToCopy = @(
    ".gitkeep",
    "rule_processor_template.py",
    "script_push_helper.py"
)

foreach ($fileName in $filesToCopy) {
    $source = Join-Path $projectRoot "scripts\$fileName"
    if (Test-Path $source) {
        Copy-Item -Force $source $scriptsDist
    }
}

Write-Host "[3/3] Verifying packaged scripts folder..." -ForegroundColor Cyan
Get-ChildItem -Force $scriptsDist | Select-Object Name, Length

Write-Host ""
Write-Host "Build complete:" -ForegroundColor Green
Write-Host "  EXE:     $distRoot\QuantumBot.exe"
Write-Host "  Scripts: $scriptsDist"
