# Lokaal testen zonder te publiceren: maakt dev\manifest.local.xml (code via https://localhost:3018)
# en registreert dat manifest in Excel. Daarna: node dev/serve.js 3018 --https  en  dev\start-excel.ps1
# Terug naar de gepubliceerde versie: scripts\registreer.ps1
param([string]$Base = 'https://localhost:3018')

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$source = Join-Path $root 'manifest.xml'
$target = Join-Path $PSScriptRoot 'manifest.local.xml'

$xml = Get-Content -LiteralPath $source -Raw -Encoding UTF8
$xml = $xml.Replace('https://kurtschaut.github.io/excel-celdetails', $Base)
$xml = $xml.Replace('<AppDomain>https://kurtschaut.github.io</AppDomain>', "<AppDomain>$Base</AppDomain>")
$xml = $xml.Replace('DefaultValue="Broncellen-preview"', 'DefaultValue="Broncellen-preview (lokaal)"')
[System.IO.File]::WriteAllText($target, $xml, (New-Object System.Text.UTF8Encoding $false))

& (Join-Path $root 'scripts\registreer.ps1') -Manifest $target
