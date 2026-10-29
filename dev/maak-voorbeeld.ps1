# Maakt dev\voorbeeld.xlsx: een testwerkmap met formules voor alle gevallen die de add-in moet aankunnen.
$ErrorActionPreference = 'Stop'
$path = Join-Path $PSScriptRoot 'voorbeeld.xlsx'
function Rgb($r, $g, $b) { $r + 256 * $g + 65536 * $b }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $res = $wb.Worksheets.Item(1); $res.Name = 'Resultaten'
  $kos = $wb.Worksheets.Item(2); $kos.Name = 'Kosten'
  $par = $wb.Worksheets.Item(3); $par.Name = 'Parameters'

  # Parameters + naam Btw
  $par.Range('A1').Value2 = 'Parameter'; $par.Range('B1').Value2 = 'Waarde'
  $par.Range('A1:B1').Font.Bold = $true
  $par.Range('A2').Value2 = 'Btw'; $par.Range('B2').Value2 = 0.21
  $par.Range('A3').Value2 = 'Korting'; $par.Range('B3').Value2 = 0.05
  $par.Range('B2:B3').NumberFormat = '0%'
  $par.Columns.Item(1).ColumnWidth = 14
  [void]$wb.Names.Add('Btw', '=Parameters!$B$2')

  # Kosten als tabel Tabel1
  $kos.Range('B2').Value2 = 'Kost'; $kos.Range('C2').Value2 = 'Bedrag'
  $kos.Range('B3').Value2 = 'Huur'; $kos.Range('C3').Value2 = 800
  $kos.Range('B4').Value2 = 'Energie'; $kos.Range('C4').Value2 = 245.3
  $kos.Range('B5').Value2 = 'Verzekering'; $kos.Range('C5').Value2 = 61
  $table = $kos.ListObjects.Add(1, $kos.Range('B2:C5'), [Type]::Missing, 1)
  $table.Name = 'Tabel1'
  $kos.Columns.Item(2).ColumnWidth = 14

  # Resultaten: titel (samengevoegd), omzetblok met opmaak, kostenblok met verborgen rij erboven
  $res.Range('A1').Value2 = 'Resultaten 2026'
  [void]$res.Range('A1:E1').Merge()
  $res.Range('A1').Font.Bold = $true
  $res.Range('A1').Font.Size = 14
  $res.Range('A1').Interior.Color = (Rgb 221 235 247)

  $res.Range('A8').Value2 = 'Omzet'; $res.Range('B8').Value2 = 'Bedrag'
  $res.Range('A8:B8').Font.Bold = $true
  $res.Range('A8:B8').Interior.Color = (Rgb 226 239 218)
  $res.Range('A9').Value2 = 'Jan'; $res.Range('B9').Value2 = 1250
  $res.Range('A10').Value2 = 'Feb'; $res.Range('B10').Value2 = 1512.5
  $res.Range('A11').Value2 = 'Mrt'; $res.Range('B11').Value2 = 1388
  $res.Range('A12').Value2 = 'Totaal'; $res.Range('B12').Formula = '=SUM(B9:B11)'
  $res.Range('A12:B12').Font.Bold = $true
  $res.Range('B9:B12').NumberFormat = '#,##0.00'
  $res.Range('A8:B12').Borders.LineStyle = 1

  $res.Range('D38').Value2 = 'Post'
  $res.Range('D38').Font.Bold = $true
  $res.Range('D39').Value2 = 100
  $res.Range('D40').Formula = '=D39*2'
  $res.Range('D41').Value2 = 300
  $res.Rows.Item(37).Hidden = $true

  $res.Range('E2').Formula2 = '=SEQUENCE(3)'

  # Testformules: label in F, formule in G
  $tests = @(
    @('Hoofdformule', '=SUM(B10,D40)*(1+Btw)+Kosten!C4+SUM(Tabel1[Bedrag])'),
    @('Ander blad', '=B12*Parameters!B3'),
    @('Samengevoegd', '=A1'),
    @('Overloop', '=SUM(E2#)'),
    @('Functienamen', '=LOG10(B9)+ATAN2(B9,B10)'),
    @('Constante', '42'),
    @('Tekst', '="A1"&B9'),
    @('Zoeken', '=VLOOKUP("Huur",Kosten!B3:C5,2,FALSE)'),
    @('Hele kolom', '=SUM(B:B)'),
    @('Geen verwijzing', '=1+2'),
    @('Fout', '=#REF!+B9'),
    @('Tien verwijzingen', '=B9+B10+B11+B12+D39+D40+D41+E2+E3+E4'),
    @('LET', '=LET(x,B9,x*2)')
  )
  $row = 3
  foreach ($t in $tests) {
    $res.Range("F$row").Value2 = $t[0]
    $res.Range("G$row").Formula2 = $t[1]
    $row++
  }
  $res.Columns.Item(6).ColumnWidth = 16
  $res.Columns.Item(7).ColumnWidth = 12
  $res.Range('C245').Formula = '=B10+D40'

  $res.Activate()
  [void]$res.Range('G3').Select()
  if (Test-Path $path) { Remove-Item $path }
  $wb.SaveAs($path, 51)
  $wb.Close($false)
  "Opgeslagen: $path"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
