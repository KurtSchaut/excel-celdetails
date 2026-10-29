# Probeer de add-in uit zonder te publiceren: alles draait op deze pc.
# Sluit eerst Excel. Dit venster blijft open zolang Excel draait; daarna wordt alles weer opgeruimd.
#   powershell -ExecutionPolicy Bypass -File dev\probeer-lokaal.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
  throw 'Sluit eerst Excel: het lokale paneel werkt enkel in een Excel dat door dit script gestart is.'
}

# Zelfondertekend certificaat voor https://localhost (enkel als bestand, niets in Windows geïnstalleerd).
$certs = Join-Path $PSScriptRoot 'certs'
if (-not (Test-Path (Join-Path $certs 'cert.pem'))) {
  New-Item -ItemType Directory -Force -Path $certs | Out-Null
  & 'C:\Program Files\Git\usr\bin\openssl.exe' req -x509 -newkey rsa:2048 -nodes -days 30 `
    -keyout (Join-Path $certs 'key.pem') -out (Join-Path $certs 'cert.pem') `
    -subj '/CN=localhost' -addext 'subjectAltName=DNS:localhost,IP:127.0.0.1' `
    -addext 'basicConstraints=critical,CA:FALSE' -addext 'extendedKeyUsage=serverAuth' `
    -addext 'keyUsage=critical,digitalSignature,keyEncipherment' 2>$null
}

& (Join-Path $PSScriptRoot 'lokaal.ps1')
$server = Start-Process node -ArgumentList "`"$(Join-Path $PSScriptRoot 'serve.js')`" 3018 --https" -WindowStyle Hidden -PassThru
try {
  & (Join-Path $PSScriptRoot 'start-excel.ps1')
  Write-Host 'Excel draait. Sluit Excel om te stoppen.'
  Start-Sleep -Seconds 5
  Wait-Process -Name EXCEL -ErrorAction SilentlyContinue
} finally {
  Stop-Process -Id $server.Id -ErrorAction SilentlyContinue
  & (Join-Path $root 'scripts\verwijder.ps1')
}
