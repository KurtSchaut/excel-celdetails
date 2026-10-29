// Test de add-in in Excel zelf: selecteert cellen in dev/voorbeeld.xlsx en controleert het taakvenster.
// Vereist: node dev/serve.js 3018 --https, dev/start-excel.ps1, en het paneel open.
// Gebruik: node dev/excel-test.js
const fs = require("fs");
const path = require("path");
const { connect } = require("./cdp.js");

const PALETTE = ["#2f6fd6", "#d23a2c", "#8a3fb3", "#25864a", "#c62f8c", "#b4661c", "#1d7fa0"];
const [BLUE, RED] = PALETTE;
const NEUTRAL = "#68707a";

let passed = 0;
let failed = 0;
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
function check(name, ok, detail) {
  if (ok) passed++;
  else {
    failed++;
    console.log(`  ✗ ${name}${detail ? `: ${detail}` : ""}`);
  }
}
const eq = (name, actual, expected) =>
  check(name, same(actual, expected), `verwacht ${JSON.stringify(expected)}, kreeg ${JSON.stringify(actual)}`);
const has = (name, text, part) => check(name, typeof text === "string" && text.includes(part), `"${part}" niet in ${JSON.stringify(text)}`);

const addrs = (s) => s.cards.map((c) => c.addr);
const allImages = (s) => s.cards.filter((c) => !c.problem).every((c) => c.image && c.imageOk);

const cases = [
  ["Resultaten", "G3", (s) => {
    eq("formule in lokale taal", s.formula, "=SOM(B10;D40)*(1+Btw)+Kosten!C4+SOM(Tabel1[Bedrag])");
    eq("kaarten", addrs(s), ["Resultaten!B10", "Resultaten!D40", "Parameters!B2", "Kosten!C4", "Kosten!C3:C5"]);
    eq("kleuren", s.cards.map((c) => c.color), [BLUE, RED, NEUTRAL, NEUTRAL, NEUTRAL]);
    eq("labels", s.cards.map((c) => c.tag), [null, null, "naam Btw", null, "tabel"]);
    eq("verwijzingen in formule", s.refs.map((r) => r.text), ["B10", "D40", "Btw", "Kosten!C4"]);
    eq("ongekleurde verwijzingen", s.refs.map((r) => r.plain), [false, false, true, true]);
    check("afbeeldingen geladen", allImages(s));
    eq("omgeving B10: kolommen", s.cards[0].cols, "ABCD");
    eq("omgeving B10: rijen", s.cards[0].rows, "8,9,10,11,12");
    has("waarde B10", s.cards[0].content, "1512,5");
    has("formule D40", s.cards[1].content, "=D39*2");
    has("resultaat D40", s.cards[1].content, "= 200");
    eq("geen melding", s.status, null);
  }],
  ["Resultaten", "G4", (s) => {
    eq("kaarten", addrs(s), ["Resultaten!B12", "Parameters!B3"]);
    eq("kleuren", s.cards.map((c) => c.color), [BLUE, NEUTRAL]);
    has("formule B12", s.cards[0].content, "=SOM(B9:B11)");
  }],
  ["Resultaten", "G5", (s) => {
    eq("samengevoegde cel", addrs(s), ["Resultaten!A1"]);
    eq("rijen", s.cards[0].rows, "1,2,3");
    check("afbeelding", allImages(s));
  }],
  ["Resultaten", "G6", (s) => {
    eq("overloopbereik", addrs(s), ["Resultaten!E2:E4"]);
    eq("label", s.cards[0].tag, "overloopbereik");
    eq("verwijzing", s.refs.map((r) => r.text), ["E2#"]);
  }],
  ["Resultaten", "G7", (s) => {
    eq("functienamen zijn geen cellen", addrs(s), ["Resultaten!B9", "Resultaten!B10"]);
    eq("zelfde verwijzing zelfde kleur", s.refs.map((r) => r.color), [BLUE, BLUE, RED]);
  }],
  ["Resultaten", "G8", (s) => {
    has("constante", s.empty, "Geen formule");
    eq("waarde getoond", s.formula, "42");
  }],
  ["Resultaten", "G9", (s) => eq("tekst genegeerd", addrs(s), ["Resultaten!B9"])],
  ["Resultaten", "G10", (s) => {
    eq("bereik op ander blad", addrs(s), ["Kosten!B3:C5"]);
    has("bereikinfo", s.cards[0].content, "Bereik van 6 cellen");
  }],
  ["Resultaten", "G11", (s) => {
    eq("hele kolom", addrs(s), ["Resultaten!B:B"]);
    has("ingekort", s.cards[0].note, "eerste 30 rijen");
    eq("30 rijen", s.cards[0].rows.split(",").length, 30);
  }],
  ["Resultaten", "G12", (s) => has("geen broncellen", s.empty, "Geen broncellen")],
  ["Resultaten", "G13", (s) => eq("foutwaarde genegeerd", addrs(s), ["Resultaten!B9"])],
  ["Resultaten", "G14", (s) => {
    eq("tien kaarten", s.cards.length, 10);
    eq("kleuren lopen rond", s.cards.map((c) => c.color), [...PALETTE, ...PALETTE.slice(0, 3)]);
    check("afbeeldingen geladen", allImages(s));
  }],
  ["Resultaten", "G15", (s) => eq("LET-variabele genegeerd", addrs(s), ["Resultaten!B9"])],
  ["Resultaten", "C245", (s) => eq("voorbeeld uit het plan", addrs(s), ["Resultaten!B10", "Resultaten!D40"])],
  ["Kosten", "C4", (s) => {
    eq("ander blad actief", s.cell, "Kosten!C4");
    has("geen formule", s.empty, "Geen formule");
  }],
];

async function main() {
  const pane = await connect();
  const run = (code) => pane.evaluate(`(async () => { ${code} })()`);
  await pane.evaluate(fs.readFileSync(path.join(__dirname, "excel-test-pagina.js"), "utf8"));
  const times = [];
  // Begin op een lege cel: een cel opnieuw selecteren geeft in Excel geen selectiegebeurtenis.
  await run(`return __t.select("Resultaten", "A20");`);

  for (const [sheet, address, verify] of cases) {
    console.log(`${sheet}!${address}`);
    const s = await run(`return __t.select(${JSON.stringify(sheet)}, ${JSON.stringify(address)});`);
    times.push(s.ms);
    eq("adres bovenaan", s.cell, `${sheet}!${address}`);
    verify(s);
  }

  console.log("Muis over een kaart");
  let s = await run(`
    await __t.select("Resultaten", "G3");
    document.querySelectorAll(".card")[1].dispatchEvent(new PointerEvent("pointerover", { bubbles: true }));
    return __t.snap();`);
  eq("verwijzing licht op", s.refs.filter((r) => r.hot).map((r) => r.text), ["D40"]);
  eq("kaart licht op", s.cards.map((c) => c.hot), [false, true, false, false, false]);
  s = await run(`
    document.getElementById("formula").querySelector(".ref").dispatchEvent(new PointerEvent("pointerover", { bubbles: true }));
    return __t.snap();`);
  eq("omgekeerd: formule → kaart", s.cards.map((c) => c.hot), [true, false, false, false, false]);

  console.log("Ga naar en terug");
  s = await run(`
    __t.mark();
    document.querySelectorAll(".card .goto")[2].click();
    await __t.rerendered();
    return __t.snap();`);
  eq("naar de naam Btw", s.cell, "Parameters!B2");
  check("terugknop actief", s.back);
  s = await run(`
    __t.mark();
    document.getElementById("back").click();
    await __t.rerendered();
    return __t.snap();`);
  eq("terug op de formulecel", s.cell, "Resultaten!G3");
  check("terugknop weer uit", !s.back);

  console.log("Pauze");
  s = await run(`
    document.getElementById("pause").click();
    __t.mark();
    await __t.excel((ctx) => ctx.workbook.worksheets.getItem("Resultaten").getRange("G4").select());
    await __t.sleep(1200);
    return __t.snap();`);
  eq("pauze negeert selectie", s.cell, "Resultaten!G3");
  s = await run(`
    document.getElementById("pause").click();
    await __t.rerendered();
    return __t.snap();`);
  eq("hervatten volgt selectie", s.cell, "Resultaten!G4");

  console.log("Instellingen");
  s = await run(`
    __t.mark();
    const r = document.getElementById("ctx-rows"), c = document.getElementById("ctx-cols");
    r.value = 0; c.value = 0; r.dispatchEvent(new Event("change"));
    await __t.rerendered();
    return __t.snap();`);
  eq("omgeving 0: enkel de broncel", [s.cards[0].cols, s.cards[0].rows], ["B", "12"]);
  s = await run(`
    __t.mark();
    const m = document.getElementById("mode"); m.value = "grid"; m.dispatchEvent(new Event("change"));
    await __t.rerendered();
    return __t.snap();`);
  check("snelle weergave toont rooster", s.cards.every((c) => c.grid && !c.image));
  await run(`
    __t.mark();
    const r = document.getElementById("ctx-rows"), c = document.getElementById("ctx-cols"), m = document.getElementById("mode");
    r.value = 2; c.value = 2; m.value = "image"; m.dispatchEvent(new Event("change"));
    await __t.rerendered();`);

  console.log("Verborgen rij in de omgeving");
  s = await run(`
    await __t.excel((ctx) => { ctx.workbook.worksheets.getItem("Resultaten").getRange("11:11").rowHidden = true; });
    const snap = await __t.select("Resultaten", "C245");
    await __t.excel((ctx) => { ctx.workbook.worksheets.getItem("Resultaten").getRange("11:11").rowHidden = false; });
    return snap;`);
  eq("rij 11 overgeslagen", s.cards[0].rows, "8,9,10,12");
  check("afbeelding met verborgen rij", allImages(s));

  pane.close();
  const sorted = [...times].sort((a, b) => a - b);
  console.log(`\nTijd per selectie (incl. 120 ms wachttijd): mediaan ${sorted[Math.floor(sorted.length / 2)]} ms, max ${sorted[sorted.length - 1]} ms`);
  console.log(`${passed} geslaagd, ${failed} mislukt`);
  process.exit(failed ? 1 : 0);
}

main().catch((e) => {
  console.error("Testfout:", e.message || e);
  process.exit(2);
});
