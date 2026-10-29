// Uitvoeren met: node --test test/
const test = require("node:test");
const assert = require("node:assert/strict");
const F = require("../formula.js");

const texts = (formula) => F.parse(formula).refs.map((r) => r.text);

test("eenvoudige som", () => {
  assert.deepEqual(texts("=B10+D40"), ["B10", "D40"]);
});

test("functienamen die op een cel lijken zijn geen verwijzing", () => {
  assert.deepEqual(texts("=LOG10(A1)+ATAN2(B2;C3)"), ["A1", "B2", "C3"]);
});

test("bladnamen, met en zonder aanhalingstekens", () => {
  const refs = F.parse("=Blad2!A1+'Mijn blad'!B2:C3+'Jan''s'!$D$4").refs;
  assert.deepEqual(refs.map((r) => [r.sheet, r.address]), [
    ["Blad2", "A1"],
    ["Mijn blad", "B2:C3"],
    ["Jan's", "$D$4"],
  ]);
  assert.equal(refs[1].text, "'Mijn blad'!B2:C3");
});

test("tekst tussen aanhalingstekens wordt genegeerd", () => {
  assert.deepEqual(texts('="A1 en ""B2"""&C3'), ["C3"]);
});

test("kolom- en rijbereiken", () => {
  assert.deepEqual(texts("=SOM(A:A;3:5;$B$2:$C$9)"), ["A:A", "3:5", "$B$2:$C$9"]);
  assert.deepEqual(F.parseArea("A:A"), { r1: 0, c1: 0, r2: F.MAX_ROW - 1, c2: 0 });
});

test("externe werkmappen worden herkend", () => {
  const refs = F.parse("=[Map2.xlsx]Blad1!A1+'C:\\map\\[Map2.xlsx]Blad 1'!B2+[1]Blad1!C3").refs;
  assert.equal(refs.length, 3);
  assert.ok(refs.every((r) => r.external));
});

test("gestructureerde verwijzingen", () => {
  const p = F.parse("=SOM(Tabel1[Bedrag])+[@Prijs]*A1");
  assert.deepEqual(p.refs.map((r) => r.text), ["A1"]);
  assert.equal(p.hasStructured, true);
  assert.deepEqual(p.names, []);
});

test("namen worden kandidaat, getallen niet", () => {
  const p = F.parse("=Btw*A1+1,5+1E+5+WAAR");
  assert.deepEqual(p.names.map((x) => x.name), ["Btw", "WAAR"]);
  assert.deepEqual(p.refs.map((r) => r.text), ["A1"]);
});

test("naam op blad", () => {
  const p = F.parse("=Blad1!Tarief*2");
  assert.deepEqual(p.names, [{ name: "Tarief", sheet: "Blad1", start: 1, end: 13 }]);
});

test("foutwaarden zijn geen bladnamen", () => {
  const p = F.parse("=#VERW!+A1+#N/B");
  assert.deepEqual(p.refs.map((r) => r.text), ["A1"]);
  assert.deepEqual(p.names, []);
});

test("overloopbereik", () => {
  const [r] = F.parse("=SOM(A1#)").refs;
  assert.equal(r.text, "A1#");
  assert.equal(r.spill, true);
  assert.equal(r.address, "A1");
});

test("3D-verwijzing", () => {
  const [r] = F.parse("=SOM(Blad1:Blad3!A1)").refs;
  assert.equal(r.threeD, true);
});

test("posities kloppen met de tekst", () => {
  const f = "=ALS(Blad2!B4>0;'Q1 2024'!C7;D9)";
  for (const r of F.parse(f).refs) assert.equal(f.slice(r.start, r.end), r.text);
});

test("ongeldige cellen zijn geen verwijzing", () => {
  assert.deepEqual(texts("=XFE1+A0+A1048577+XFD1048576"), ["XFD1048576"]);
});

test("adressen omzetten", () => {
  assert.equal(F.areaToA1(F.parseArea("$C$3:$A$1")), "A1:C3");
  assert.equal(F.areaToA1(F.parseArea("B:D")), "B:D");
  assert.equal(F.areaToA1(F.parseArea("4:4")), "4:4");
  assert.equal(F.indexToCol(0), "A");
  assert.equal(F.indexToCol(26), "AA");
  assert.equal(F.indexToCol(16383), "XFD");
  assert.equal(F.colToIndex("XFD"), 16383);
});

test("adreslijsten splitsen", () => {
  assert.deepEqual(F.splitAddressList("Blad1!A1:B4, 'Mijn, blad'!D1:D4, E5"), [
    { sheet: "Blad1", address: "A1:B4" },
    { sheet: "Mijn, blad", address: "D1:D4" },
    { sheet: "Mijn, blad", address: "E5" },
  ]);
});

test("bladnamen citeren", () => {
  assert.equal(F.quoteSheet("Blad1"), "Blad1");
  assert.equal(F.quoteSheet("Mijn blad"), "'Mijn blad'");
  assert.equal(F.quoteSheet("A1"), "'A1'");
  assert.equal(F.quoteSheet("2024"), "'2024'");
  assert.equal(F.quoteSheet("Jan's"), "'Jan''s'");
});

test("kleuren: zelfde verwijzing zelfde kleur, andere bladen ongekleurd", () => {
  const colors = F.assignColors(["a!B1", null, "a!C1", "a!B1", "a!D1"], 2);
  assert.deepEqual([...colors], [["a!B1", 0], ["a!C1", 1], ["a!D1", 0]]);
});
