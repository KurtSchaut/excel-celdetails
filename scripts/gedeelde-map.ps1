#Requires -RunAsAdministrator
# Eenmalige installatie via een gedeelde map, zoals Microsoft dat voor Excel op Windows voorziet.
# Voer dit zelf uit in PowerShell als administrator (met je eigen Windows-account):
#   powershell -ExecutionPolicy Bypass -File scripts\gedeelde-map.ps1
# Wat het doet:
#  1. maakt de map C:\Broncellen-manifest en zet manifest.xml erin;
#  2. deelt die map als \\<computernaam>\Broncellen, alleen-lezen en enkel voor jou;
#  3. zet die share in Excel bij de vertrouwde invoegtoepassingscatalogi (Vertrouwenscentrum), met "Weergeven in menu".
# Daarna in Excel: Start > Invoegtoepassingen > Meer invoegtoepassingen > Gedeelde map > Broncellen-preview > Toevoegen.
# Ongedaan maken: scripts\gedeelde-map.ps1 -Verwijder
param([string]$Map = 'C:\Broncellen-manifest', [string]$Share = 'Broncellen', [switch]$Verwijder)

$ErrorActionPreference = 'Stop'
$catalogId = '{5b0c1e4a-2f63-4c8e-9d1a-7e3b6c2f9a41}'
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"

if ($Verwijder) {
  if (Test-Path $key) { Remove-Item $key }
  if (Get-SmbShare -Name $Share -ErrorAction SilentlyContinue) { Remove-SmbShare -Name $Share -Force }
  Write-Host "Catalogus en share verwijderd. De map $Map kun je zelf weggooien."
  return
}

New-Item -ItemType Directory -Force -Path $Map | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot '..\manifest.xml') -Destination (Join-Path $Map 'manifest.xml') -Force
if (-not (Get-SmbShare -Name $Share -ErrorAction SilentlyContinue)) {
  New-SmbShare -Name $Share -Path $Map -ReadAccess "$env:USERDOMAIN\$env:USERNAME" | Out-Null
}

$url = "\\$env:COMPUTERNAME\$Share"
if (-not (Test-Path $key)) { New-Item -Path $key | Out-Null }
New-ItemProperty -Path $key -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name Url -Value $url -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
Write-Host "Klaar: $url staat als vertrouwde catalogus in Excel. Start Excel opnieuw en voeg de add-in toe via Gedeelde map."
