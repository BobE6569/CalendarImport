$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"

function Find-Python {
    $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        return @("py", "-3")
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return @("python")
    }

    throw "Python 3 was not found. Install Python 3.10 or newer, then run this script again."
}

if (-not (Test-Path $venvPython)) {
    $pythonCommand = Find-Python
    if ($pythonCommand.Length -gt 1) {
        & $pythonCommand[0] $pythonCommand[1..($pythonCommand.Length - 1)] -m venv (Join-Path $projectRoot ".venv")
    } else {
        & $pythonCommand[0] -m venv (Join-Path $projectRoot ".venv")
    }
}

& $venvPython -m pip install -r (Join-Path $projectRoot "requirements.txt")
& $venvPython -m pip install -e $projectRoot
& $venvPython -m calendar_import
