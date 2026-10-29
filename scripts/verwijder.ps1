# Haalt de add-in weer uit Excel (tegenhanger van registreer.ps1).
# Daarna Excel (opnieuw) starten; de knop "Broncellen" is dan weg.
param([string]$Manifest = (Join-Path $PSScriptRoot '..\manifest.xml'))

$ErrorActionPreference = 'Stop'
$path = (Resolve-Path -LiteralPath $Manifest).Path
$id = ([xml](Get-Content -LiteralPath $path -Raw)).OfficeApp.Id
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

if ((Test-Path $key) -and (Get-ItemProperty -Path $key -Name $id -ErrorAction SilentlyContinue)) {
  Remove-ItemProperty -Path $key -Name $id
  Write-Host "Verwijderd: $id"
} else {
  Write-Host 'De add-in was niet geregistreerd.'
}
