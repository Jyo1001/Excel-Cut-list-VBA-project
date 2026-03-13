param()

$ErrorActionPreference = "Stop"

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildLog = Join-Path $ScriptRoot "build.log"
$ValidationJson = Join-Path $ScriptRoot "ROUND_BAR_Nesting_Template.validation.json"
$PythonScript = Join-Path $ScriptRoot "build_xlsm.py"

if (Test-Path $BuildLog) {
    Remove-Item $BuildLog -Force
}
New-Item -ItemType File -Path $BuildLog -Force | Out-Null

function Write-Log {
    param(
        [string]$Level,
        [string]$Message
    )

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts | $Level | $Message"
    Write-Host $line
    Add-Content -Path $BuildLog -Value $line
}

function Get-PythonExe {
    $cmd = Get-Command py -ErrorAction SilentlyContinue
    if ($cmd) { return "py" }

    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { return "python" }

    throw "Python was not found on PATH."
}

function Install-PyWin32IfMissing {
    param([string]$PyCmd)

    Write-Log "INFO " "Checking pywin32..."

    & $PyCmd -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('win32com.client') else 1)"
    $exitCode = $LASTEXITCODE

    if ($exitCode -eq 0) {
        Write-Log "INFO " "pywin32 already available."
        return
    }

    Write-Log "INFO " "Installing pywin32..."
    & $PyCmd -m pip install pywin32 2>&1 | ForEach-Object {
        $line = $_.ToString()
        Write-Host $line
        Add-Content -Path $BuildLog -Value $line
    }

    if ($LASTEXITCODE -ne 0) {
        throw "pywin32 installation failed."
    }

    Write-Log "INFO " "pywin32 install complete."
}

function Close-ExcelProcesses {
    Write-Log "INFO " "Closing Excel processes..."
    Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $_.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 400
            if (-not $_.HasExited) {
                $_.Kill()
            }
        }
        catch {
            Write-Log "WARN " "Could not close EXCEL PID=$($_.Id): $($_.Exception.Message)"
        }
    }
}

try {
    Write-Log "INFO " "ROUND BAR XLSM builder starting..."
    Write-Log "INFO " "Working folder: $ScriptRoot"

    Close-ExcelProcesses

    $PyCmd = Get-PythonExe
    Write-Log "INFO " "Python command: $PyCmd"

    Install-PyWin32IfMissing -PyCmd $PyCmd

    if (-not (Test-Path $PythonScript)) {
        throw "build_xlsm.py not found: $PythonScript"
    }

    Write-Log "INFO " "Running build_xlsm.py..."

    & $PyCmd $PythonScript 2>&1 | ForEach-Object {
        $line = $_.ToString()
        Write-Host $line
        Add-Content -Path $BuildLog -Value $line
    }

    $pythonExit = $LASTEXITCODE

    if ($pythonExit -ne 0) {
        throw "build_xlsm.py returned exit code $pythonExit"
    }

    Write-Log "INFO " "Build completed successfully."

    if (Test-Path $ValidationJson) {
        Write-Log "INFO " "Validation JSON: $ValidationJson"
    }
}
catch {
    Write-Log "ERROR" $_.Exception.Message
    Write-Log "WARN " "Showing last 120 lines of build.log..."

    if (Test-Path $BuildLog) {
        Get-Content $BuildLog -Tail 120 | Out-Host
    }

    if (Test-Path $ValidationJson) {
        Write-Log "WARN " "Check validation JSON: $ValidationJson"
    }
    else {
        Write-Log "WARN " "Validation JSON not found."
    }

    exit 1
}