# Broncellen-preview

Excel-invoegtoepassing (Office-add-in) voor Excel op Windows. Selecteer je een cel met een formule,
dan toont een zijpaneel elke broncel waarnaar die formule verwijst, met de omliggende cellen erbij.
Je hoeft dus niet meer terug te scrollen.

Per broncel krijg je een kaart met:

- het adres (`Blad!Cel`), in dezelfde kleur als Excel de verwijzing kleurt tijdens het bewerken;
- een uitsnede met kolom- en rijkoppen en de echte opmaak, waarin de broncel gemarkeerd is;
- de formule (of waarde) van de broncel en het resultaat;
- een knop **Ga naar**. Met de pijl linksboven ga je terug naar de vorige cel.

Ga je met de muis over een kaart, dan licht de bijhorende verwijzing op in de formule bovenaan het
paneel, en omgekeerd. Klik je op een verwijzing in die formule, dan scrolt het paneel naar de kaart.

## Bediening

| Wat | Hoe |
| --- | --- |
| Paneel openen | Start > **Broncellen** (lintknop) |
| Paneel sluiten | het kruisje rechtsboven in het paneel |
| Tijdelijk niet reageren op selecties | de pauzeknop naast de vernieuwknop (groen = gepauzeerd) |
| Aantal rijen/kolommen rond de broncel | **Rijen ±** en **Kolommen ±** (standaard 2) |
| Weergave | **Opmaak**: Excel tekent de uitsnede met de echte opmaak. **Snel**: eenvoudig rooster met tekst, vulkleur en vet |

## Installeren en verwijderen

1. **Publiceren.** De code van het paneel moet op een https-adres staan dat Windows vertrouwt. Het
   manifest wijst naar GitHub Pages (`https://kurtschaut.github.io/excel-celdetails/`).
2. **Gedeelde map instellen** (eenmalig, in een gewone PowerShell; Windows vraagt zelf om
   beheerdersrechten voor de share):

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\gedeelde-map.ps1
   ```

   Dit deelt een map met `manifest.xml` als `\\<computernaam>\Broncellen` en zet die bij de
   vertrouwde invoegtoepassingscatalogi van Excel.
3. **Toevoegen in Excel:** Start > Invoegtoepassingen > Meer invoegtoepassingen > **Gedeelde map** >
   Broncellen-preview > Toevoegen. Vanaf dan staat de knop op het tabblad Start.

Verwijderen: in Excel via Start > Invoegtoepassingen (de add-in verwijderen), daarna
`scripts\gedeelde-map.ps1 -Verwijder`.

`scripts\registreer.ps1` meldt het manifest aan als ontwikkelaars-add-in. In Excel 2024 (niet-365)
geeft dat geen blijvende knop: die verschijnt enkel zolang een werkmap open is die het paneel bevat,
zoals `dev\voorbeeld.xlsx`. Gebruik het daarom alleen om te testen.

### Uitproberen zonder te publiceren

```powershell
powershell -ExecutionPolicy Bypass -File dev\probeer-lokaal.ps1
```

Sluit eerst Excel. Het script start een lokale https-server en opent Excel met `dev\voorbeeld.xlsx`,
waarin het paneel al gekoppeld is. Het lokale certificaat wordt nergens in Windows geïnstalleerd:
enkel dat ene Excel-proces aanvaardt die ene sleutel. Sluit je Excel, dan ruimt het script alles op.

## Wat werkt niet

- Een zwevende popup naast de cel: Office-add-ins kunnen geen venster aan een cel vastmaken.
- Reageren op hoveren over de formulebalk zelf: dat kan een add-in niet zien. Het paneel volgt de geselecteerde cel.
- Verwijzingen naar andere werkmappen en 3D-verwijzingen (`Blad1:Blad3!A1`). Ze krijgen een kaart met uitleg.
- Dynamische verwijzingen via `INDIRECT` of `VERSCHUIVING`.

Verwijzingen naar andere bladen, gedefinieerde namen, overloopbereiken (`A1#`) en tabelverwijzingen
(`Tabel1[Kolom]`) werken wel. Die laatste laat de add-in door Excel zelf oplossen.

## Vertrouwelijkheid

Enkel de programmacode staat online. De add-in leest de werkmap in Excel zelf, en de uitsneden
worden in het paneel getekend. Er gaan geen celgegevens naar een server.

## Ontwikkelen en testen

Er is geen buildstap: `taskpane.html`, `taskpane.css`, `taskpane.js` en `formula.js` worden zo gepubliceerd.

```powershell
node --test test/formula.test.js   # tests van de formule-analyse
node dev/serve.js                  # voorbeeld met nepgegevens op http://localhost:3017
```

Test in Excel zelf (Excel gestart via `dev\probeer-lokaal.ps1`, paneel open):

```powershell
node dev/excel-test.js   # selecteert de testcellen in dev\voorbeeld.xlsx en controleert het paneel
node dev/cdp.js "<javascript>"   # voert code uit in het paneel; --watch 10 toont de console
```

`dev\maak-voorbeeld.ps1` maakt de testwerkmap opnieuw aan; `python -I dev\koppel-paneel.py dev\voorbeeld.xlsx`
koppelt er daarna het paneel aan.

Excel bewaart de bestanden van een add-in in een cache. Verhoog na een wijziging het `?v=`-nummer in
`taskpane.html`. Helpt dat niet, sluit Excel en leeg `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
