# Meldt de add-in aan als ontwikkelaars-add-in, enkel voor de huidige Windows-gebruiker (om te testen).
# Let op: in Excel 2024 (niet-365) geeft dit geen blijvende lintknop. De knop verschijnt enkel zolang een
# werkmap open is die het paneel bevat, zoals dev\voorbeeld.xlsx. Voor gewoon gebruik: scripts\gedeelde-map.ps1.
param([string]$Manifest = (Join-Path $PSScriptRoot '..\manifest.xml'))

$ErrorActionPreference = 'Stop'
$path = (Resolve-Path -LiteralPath $Manifest).Path
$id = ([xml](Get-Content -LiteralPath $path -Raw)).OfficeApp.Id
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

if (-not (Test-Path $key)) { New-Item -Path $key | Out-Null }
New-ItemProperty -Path $key -Name $id -Value $path -PropertyType String -Force | Out-Null
Write-Host "Geregistreerd: $id -> $path"
