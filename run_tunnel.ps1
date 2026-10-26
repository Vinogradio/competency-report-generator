$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$desiredUrl = "https://otchet-kompetencii-a50c.loca.lt"
$logDir = Join-Path $env:TEMP "opencode"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$outLog = Join-Path $logDir "localtunnel.out.log"
$errLog = Join-Path $logDir "localtunnel.err.log"

$npx = (Get-Command npx.cmd -ErrorAction SilentlyContinue).Source
if (-not $npx) {
    "npx.cmd not found" | Set-Content -Path $errLog -Encoding UTF8
    exit 1
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Stop-Localtunnel {
    $processes = Get-CimInstance Win32_Process | Where-Object {
        ($_.Name -eq "node.exe" -and $_.CommandLine -match "localtunnel|lt\.js") -or
        ($_.Name -eq "cmd.exe" -and $_.CommandLine -match "localtunnel")
    }
    foreach ($process in $processes) {
        Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

function Test-Tunnel {
    $request = [System.Net.HttpWebRequest]::Create($desiredUrl + "/")
    $request.Method = "HEAD"
    $request.Timeout = 15000
    $request.AllowAutoRedirect = $true
    try {
        $response = $request.GetResponse()
        $code = [int]$response.StatusCode
        $response.Close()
        return ($code -eq 200 -or $code -eq 401)
    } catch [System.Net.WebException] {
        if ($_.Exception.Response) {
            $code = [int]$_.Exception.Response.StatusCode
            return ($code -eq 200 -or $code -eq 401)
        }
        return $false
    } catch {
        return $false
    }
}

while ($true) {
    Stop-Localtunnel
    Remove-Item $outLog, $errLog -ErrorAction SilentlyContinue

    $process = Start-Process -FilePath $npx -ArgumentList @(
        "--yes",
        "localtunnel",
        "--port", "5000",
        "--local-host", "127.0.0.1",
        "--subdomain", "otchet-kompetencii-a50c"
    ) -RedirectStandardOutput $outLog -RedirectStandardError $errLog -WindowStyle Hidden -PassThru

    $url = $null
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Seconds 1
        if (Test-Path $outLog) {
            $match = Select-String -Path $outLog -Pattern "https://[a-z0-9-]+\.loca\.lt" | Select-Object -Last 1
            if ($match) {
                $url = $match.Matches[0].Value
                break
            }
        }
    }

    if ($url -ne $desiredUrl) {
        taskkill /PID $process.Id /T /F 2>$null | Out-Null
        Stop-Localtunnel
        Start-Sleep -Seconds 30
        continue
    }

    while (-not $process.HasExited) {
        Start-Sleep -Seconds 30
        if (-not (Test-Tunnel)) {
            taskkill /PID $process.Id /T /F 2>$null | Out-Null
            Stop-Localtunnel
            break
        }
    }

    Start-Sleep -Seconds 30
}
