$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$lockPath = Join-Path $env:LOCALAPPDATA "notaria_sofia_ode\creditos_web_local\app.lock"
$dir = Split-Path -Parent $lockPath
New-Item -ItemType Directory -Force -Path $dir | Out-Null
if (Test-Path $lockPath) {
    Write-Host "La aplicación ya está ejecutándose o existe un bloqueo previo."
    exit 1
}
"running" | Set-Content -Path $lockPath
try {
    $env:PYTHONPATH = $root
    Start-Process "http://127.0.0.1:8000/"
    python -m uvicorn app:app --host 127.0.0.1 --port 8000
} finally {
    Remove-Item -Force $lockPath -ErrorAction SilentlyContinue
}
