$ErrorActionPreference = "Stop"

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptRoot

$Builder = Join-Path $ScriptRoot "build_xlsm.ps1"

if (-not (Test-Path $Builder)) {
    Write-Host "ERROR: build_xlsm.ps1 not found in $ScriptRoot" -ForegroundColor Red
    exit 1
}

Write-Host "Running builder from: $ScriptRoot"
powershell -NoProfile -ExecutionPolicy Bypass -File $Builder
exit $LASTEXITCODE