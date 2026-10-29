# Start Excel met de testwerkmap, voor lokaal testen via https://localhost:3018 (node dev/serve.js 3018 --https).
# Enkel voor dit Excel-proces krijgt WebView2:
#  - een debugpoort, zodat dev/cdp.js in het taakvenster kan kijken;
#  - een uitzondering voor precies één certificaatsleutel: het zelfondertekende certificaat in dev\certs.
# Er wordt niets in Windows geïnstalleerd of blijvend ingesteld.
param([string]$Workbook = (Join-Path $PSScriptRoot 'voorbeeld.xlsx'), [int]$Port = 9222)

$ssl = 'C:\Program Files\Git\usr\bin\openssl.exe'
$cert = Join-Path $PSScriptRoot 'certs\cert.pem'
# De PEM-publieke sleutel is de SubjectPublicKeyInfo; Chromium verwacht de SHA-256 daarvan in base64.
$pub = ((& $ssl x509 -in $cert -pubkey -noout) -join '') -replace '-----[^-]+-----', ''
$spki = [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([Convert]::FromBase64String($pub)))

$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Port --ignore-certificate-errors-spki-list=$spki"
Write-Host "WebView2: $env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"
Start-Process excel.exe -ArgumentList "`"$Workbook`""
