# Sunday Service slides app - Windows setup & run (PowerShell).
#
#   .\setup.ps1 install   create .venv and install dependencies
#   .\setup.ps1 run       install (if needed) then launch the app
#   .\setup.ps1 verify    check the Google Sheets connection
#   .\setup.ps1 stop      stop a running app
#   .\setup.ps1 clean     remove .venv
#
# If scripts are blocked, run once:
#   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
# or launch via the .bat wrappers (install.bat / run.bat).

param(
    [ValidateSet('install', 'run', 'verify', 'stop', 'clean', 'help')]
    [string]$Task = 'install',
    [int]$Port = 8501
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

$Venv      = '.venv'
$Py        = Join-Path $Venv 'Scripts\python.exe'
$Pip       = Join-Path $Venv 'Scripts\pip.exe'
$Streamlit = Join-Path $Venv 'Scripts\streamlit.exe'

function Initialize-Venv {
    if (Test-Path $Py) { return }
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $Venv
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv $Venv
    } else {
        throw 'Python 3 not found. Install from https://www.python.org/downloads/windows/'
    }
}

function Install-App {
    Initialize-Venv
    & $Pip install --upgrade pip
    & $Pip install -r requirements.txt
}

function Show-Help {
    Write-Host @'
Usage: .\setup.ps1 <task>

  install   Create .venv and install dependencies
  run       Launch the app locally (http://localhost:8501)
  verify    Check the Google Sheets connection
  stop      Stop a running app
  clean     Remove .venv
  help      Show this message
'@
}

switch ($Task) {
    'install' { Install-App }
    'run' {
        Install-App
        & $Streamlit run app.py --server.port $Port
    }
    'verify' {
        Install-App
        & $Py -c "import tomllib, store; s = store.get_store(tomllib.load(open('.streamlit/secrets.toml','rb'))); print('backend:', type(s).__name__); print('weeks:', list(s.all().keys()))"
    }
    'stop' {
        Get-CimInstance Win32_Process -Filter "Name = 'python.exe'" |
            Where-Object { $_.CommandLine -like '*streamlit*' } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
        Write-Host 'Stopped.'
    }
    'clean' { Remove-Item -Recurse -Force $Venv -ErrorAction SilentlyContinue }
    'help'  { Show-Help }
}
