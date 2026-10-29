# Eenmalige installatie via een gedeelde map, zoals Microsoft dat voor Excel op Windows voorziet.
# Voer dit uit in een gewone PowerShell; Windows vraagt zelf om beheerdersrechten voor de share:
#   powershell -ExecutionPolicy Bypass -File "C:\Users\L480\Documents\GitHub\excel-celdetails\scripts\gedeelde-map.ps1"
# Wat het doet:
#  1. maakt de map C:\Broncellen-manifest en zet manifest.xml erin (met beheerdersrechten);
#  2. deelt die map als \\<computernaam>\Broncellen, alleen-lezen en enkel voor jou (met beheerdersrechten);
#  3. zet die share in Excel bij de vertrouwde invoegtoepassingscatalogi (Vertrouwenscentrum), met "Weergeven in menu".
# Daarna in Excel: Start > Invoegtoepassingen > Meer invoegtoepassingen > Gedeelde map > Broncellen-preview > Toevoegen.
# Ongedaan maken: zelfde commando met -Verwijder
param(
  [string]$Map = 'C:\Broncellen-manifest',
  [string]$Share = 'Broncellen',
  [string]$Gebruiker = "$env:USERDOMAIN\$env:USERNAME",
  [switch]$Verwijder,
  [switch]$AlleenShare
)

$ErrorActionPreference = 'Stop'
$catalogId = '{5b0c1e4a-2f63-4c8e-9d1a-7e3b6c2f9a41}'
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
  [Security.Principal.WindowsBuiltInRole]::Administrator)

# Map en share: dit deel draait met beheerdersrechten.
if ($AlleenShare) {
  if ($Verwijder) {
    if (Get-SmbShare -Name $Share -ErrorAction SilentlyContinue) { Remove-SmbShare -Name $Share -Force }
    return
  }
  New-Item -ItemType Directory -Force -Path $Map | Out-Null
  Copy-Item -LiteralPath (Join-Path $PSScriptRoot '..\manifest.xml') -Destination (Join-Path $Map 'manifest.xml') -Force
  if (-not (Get-SmbShare -Name $Share -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name $Share -Path $Map -ReadAccess $Gebruiker | Out-Null
  }
  return
}

$shareArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-AlleenShare',
  '-Map', "`"$Map`"", '-Share', $Share, '-Gebruiker', "`"$Gebruiker`"")
if ($Verwijder) { $shareArgs += '-Verwijder' }
if ($isAdmin) { & $PSCommandPath -AlleenShare -Map $Map -Share $Share -Gebruiker $Gebruiker -Verwijder:$Verwijder }
else {
  Write-Host 'Windows vraagt nu om beheerdersrechten voor de gedeelde map: kies Ja.'
  Start-Process powershell -Verb RunAs -Wait -ArgumentList $shareArgs
}

# Vertrouwde catalogus: dit deel draait als jezelf, zodat het in jouw Excel-instellingen komt.
if ($Verwijder) {
  if (Test-Path $key) { Remove-Item $key }
  Write-Host "Catalogus en share verwijderd. De map $Map kun je zelf weggooien."
  return
}
if (-not (Get-SmbShare -Name $Share -ErrorAction SilentlyContinue)) {
  throw 'De gedeelde map is niet aangemaakt (beheerdersrechten geweigerd?). Er is niets gewijzigd in Excel.'
}
$url = "\\$env:COMPUTERNAME\$Share"
if (-not (Test-Path $key)) { New-Item -Path $key | Out-Null }
New-ItemProperty -Path $key -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name Url -Value $url -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
Write-Host "Klaar: $url staat als vertrouwde catalogus in Excel. Start Excel opnieuw en voeg de add-in toe via Gedeelde map."
