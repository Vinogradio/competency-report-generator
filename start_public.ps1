$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"
if (-not (Test-Path $python)) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCommand) {
        Write-Host "Python not found."
        exit 1
    }
    $python = $pythonCommand.Source
}

$listener = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue
if (-not $listener) {
    Write-Host "Starting the app..."
    Start-Process -FilePath $python -ArgumentList "app.py" -WorkingDirectory $root -WindowStyle Hidden
    Start-Sleep -Seconds 4
    $listener = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue
    if (-not $listener) {
        Write-Host "The app did not start on port 5000."
        exit 1
    }
}

$oldSupervisors = Get-CimInstance Win32_Process | Where-Object {
    $_.Name -eq "powershell.exe" -and $_.CommandLine -like "*run_tunnel.ps1*"
}
foreach ($process in $oldSupervisors) {
    Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
}

$oldTunnels = Get-CimInstance Win32_Process | Where-Object {
    ($_.Name -eq "ssh.exe" -and $_.CommandLine -match "localhost\.run|serveo\.net") -or
    ($_.Name -eq "node.exe" -and $_.CommandLine -match "localtunnel|lt\.js") -or
    ($_.Name -eq "cmd.exe" -and $_.CommandLine -match "localtunnel")
}
foreach ($process in $oldTunnels) {
    Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
}
Start-Sleep -Seconds 1

$logDir = Join-Path $env:TEMP "opencode"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$outLog = Join-Path $logDir "localtunnel.out.log"
$errLog = Join-Path $logDir "localtunnel.err.log"
Remove-Item $outLog, $errLog -ErrorAction SilentlyContinue

$powershell = (Get-Command powershell.exe).Source
$supervisor = Join-Path $root "run_tunnel.ps1"
Start-Process -FilePath $powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$supervisor`"" -WindowStyle Hidden | Out-Null
Write-Host "Starting the tunnel..."

$desiredUrl = "https://otchet-kompetencii-a50c.loca.lt"
$url = $null
for ($i = 0; $i -lt 90; $i++) {
    Start-Sleep -Seconds 1
    if (Test-Path $outLog) {
        $match = Select-String -Path $outLog -Pattern "https://[a-z0-9-]+\.loca\.lt" | Select-Object -Last 1
        if ($match -and $match.Matches[0].Value -eq $desiredUrl) {
            $url = $match.Matches[0].Value
            break
        }
    }
}

if (-not $url) {
    Write-Host "The tunnel did not start."
    if (Test-Path $errLog) {
        Get-Content $errLog -Tail 20
    }
    exit 1
}

$url | Set-Content -Path (Join-Path $root "public_url.txt") -Encoding ASCII
Write-Host ""
Write-Host "PUBLIC URL: $url"
Write-Host "Login: team"
Write-Host "Password: see APP_PASSWORD in .env"
Write-Host ""
Write-Host "The URL stays the same when the tunnel reconnects. Keep this computer awake."
