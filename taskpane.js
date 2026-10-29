/* Broncellen-preview — taakvenster.
   Volgt de selectie in Excel en toont per broncel van de formule een uitsnede met de omliggende cellen. */
"use strict";

// Kleuren van verwijzingen in de formulebalk tijdens het bewerken, in volgorde van voorkomen
// (nagemeten in Excel: na de zevende kleur begint de reeks opnieuw).
const REF_COLORS = ["#2f6fd6", "#d23a2c", "#8a3fb3", "#25864a", "#c62f8c", "#b4661c", "#1d7fa0"];
// Verwijzingen naar andere bladen en namen kleurt Excel niet.
const NEUTRAL = "#68707a";
const MAX_CARDS = 20;
const MAX_NAMES = 15;
const MAX_AREA_ROWS = 30;
const MAX_AREA_COLS = 12;
const PT_TO_PX = 96 / 72;
const SETTINGS_KEY = "broncellen.instellingen";
const ICON_GOTO = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M7 3H3v10h10V9M9.5 2.5h4v4M13.5 2.5 7.5 8.5"/></svg>';

const F = window.BronFormula;

const state = {
  paused: false,
  rows: 2,
  cols: 2,
  mode: "image",
  history: [],
  current: null,
  seq: 0,
  timer: 0,
  hotKey: null,
};

const $ = (id) => document.getElementById(id);
const sum = (list) => list.reduce((a, b) => a + b, 0);
const pct = (part, whole) => `${(part / whole) * 100}%`;
const clamp = (n, lo, hi) => Math.min(hi, Math.max(lo, n));
const apiSupported = (version) => Office.context.requirements.isSetSupported("ExcelApi", version);

function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === null || v === undefined || v === false) continue;
    if (k === "class") el.className = v;
    else if (k === "html") el.innerHTML = v;
    else if (k === "style") {
      for (const [prop, value] of Object.entries(v)) {
        if (value === null || value === undefined) continue;
        if (prop.startsWith("--")) el.style.setProperty(prop, value);
        else el.style[prop] = value;
      }
    } else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const child of children.flat()) {
    if (child === null || child === undefined || child === false) continue;
    el.append(child.nodeType ? child : String(child));
  }
  return el;
}

function sheetLabel(sheet, address) {
  return `${F.quoteSheet(sheet)}!${address}`;
}

// Formuletekst met breekpunten na operatoren, zodat lange formules niet midden in een woord afbreken.
function breakable(text) {
  const frag = document.createDocumentFragment();
  for (const part of text.split(/(?<=[+\-*/^&=<>;,(])/)) {
    if (frag.childNodes.length) frag.append(h("wbr"));
    frag.append(part);
  }
  return frag;
}

/* ---------- Instellingen en bediening ---------- */

function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
    if (Number.isInteger(saved.rows)) state.rows = clamp(saved.rows, 0, 10);
    if (Number.isInteger(saved.cols)) state.cols = clamp(saved.cols, 0, 10);
    if (saved.mode === "image" || saved.mode === "grid") state.mode = saved.mode;
  } catch {
    // Geen opslag beschikbaar: standaardwaarden blijven.
  }
  $("ctx-rows").value = state.rows;
  $("ctx-cols").value = state.cols;
  $("mode").value = state.mode;
}

function saveSettings() {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ rows: state.rows, cols: state.cols, mode: state.mode }));
  } catch {
    // Niet bewaard; geldt dan enkel voor deze sessie.
  }
}

function onSettingsChange() {
  const rows = parseInt($("ctx-rows").value, 10);
  const cols = parseInt($("ctx-cols").value, 10);
  state.rows = clamp(Number.isNaN(rows) ? 2 : rows, 0, 10);
  state.cols = clamp(Number.isNaN(cols) ? 2 : cols, 0, 10);
  state.mode = $("mode").value === "grid" ? "grid" : "image";
  $("ctx-rows").value = state.rows;
  $("ctx-cols").value = state.cols;
  saveSettings();
  refresh(true).catch(showError);
}

function bindUi() {
  $("pause").addEventListener("click", togglePause);
  $("resume").addEventListener("click", togglePause);
  $("refresh").addEventListener("click", () => refresh(true).catch(showError));
  $("back").addEventListener("click", goBack);
  $("ctx-rows").addEventListener("change", onSettingsChange);
  $("ctx-cols").addEventListener("change", onSettingsChange);
  $("mode").addEventListener("change", onSettingsChange);
  $("formula").addEventListener("click", onFormulaClick);
  document.addEventListener("pointerover", (e) => {
    const el = e.target.closest ? e.target.closest("[data-key]") : null;
    setHot(el ? el.dataset.key : null);
  });
  document.documentElement.addEventListener("pointerleave", () => setHot(null));
}

function togglePause() {
  state.paused = !state.paused;
  state.seq++;
  const btn = $("pause");
  btn.setAttribute("aria-pressed", String(state.paused));
  btn.title = state.paused ? "Hervatten: de selectie weer volgen" : "Pauzeren: de selectie tijdelijk niet volgen";
  btn.setAttribute("aria-label", state.paused ? "Hervatten" : "Pauze");
  document.body.classList.toggle("paused", state.paused);
  if (!state.paused) scheduleRefresh(0);
}

// Licht een verwijzing op in de formule én de bijhorende kaart.
function setHot(key) {
  if (key === state.hotKey) return;
  state.hotKey = key;
  for (const el of document.querySelectorAll(".hot")) el.classList.remove("hot");
  if (key === null) return;
  for (const el of document.querySelectorAll("[data-key]")) {
    if (el.dataset.key === key) el.classList.add("hot");
  }
}

function onFormulaClick(e) {
  const ref = e.target.closest("[data-key]");
  if (!ref) return;
  const card = [...document.querySelectorAll(".card")].find((c) => c.dataset.key === ref.dataset.key);
  if (card) card.scrollIntoView({ block: "nearest", behavior: "smooth" });
}

function updateBack() {
  const btn = $("back");
  const prev = state.history[state.history.length - 1];
  btn.disabled = !prev;
  btn.title = prev ? `Terug naar ${sheetLabel(prev.sheet, prev.address)}` : "Terug naar de vorige cel";
}

async function goTo(sheet, address, remember = true) {
  if (remember && state.current) {
    state.history.push(state.current);
    if (state.history.length > 30) state.history.shift();
  }
  updateBack();
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    ws.activate();
    ws.getRange(address).select();
    await ctx.sync();
  });
}

function goBack() {
  const prev = state.history.pop();
  updateBack();
  if (prev) goTo(prev.sheet, prev.address, false).catch(showError);
}

function setStatus(message, kind) {
  const el = $("status");
  el.hidden = !message;
  el.textContent = message || "";
  el.className = `status${kind ? ` ${kind}` : ""}`;
}

function showError(e) {
  console.error(e);
  const message = (e && (e.message || e.code)) || String(e);
  setStatus(`Er ging iets mis: ${message}`, "error");
}

/* ---------- Gegevens ophalen uit Excel ---------- */

function scheduleRefresh(delay = 120) {
  clearTimeout(state.timer);
  state.timer = setTimeout(() => refresh().catch(showError), delay);
}

async function refresh(force = false) {
  if (!force && state.paused) return;
  const seq = ++state.seq;
  const busy = setTimeout(() => document.body.classList.add("busy"), 150);
  try {
    const model = await Excel.run((ctx) => collect(ctx, seq));
    if (model && seq === state.seq) render(model);
  } finally {
    clearTimeout(busy);
    if (seq === state.seq) document.body.classList.remove("busy");
  }
}

function sourceKey(sheet, area) {
  return `${sheet.toLowerCase()}!${F.areaToA1(area)}`;
}

/* Bouwt een gewoon model (zonder Excel-objecten) van de actieve cel en haar broncellen.
   Geeft null terug als intussen een nieuwere selectie binnenkwam. */
async function collect(ctx, seq) {
  const stale = () => seq !== state.seq;
  const wb = ctx.workbook;
  const cell = wb.getActiveCell();
  cell.load("address, formulas, formulasLocal, text");
  const sheet = wb.worksheets.getActiveWorksheet();
  sheet.load("name");
  wb.worksheets.load("items/name");
  await ctx.sync();
  if (stale()) return null;

  const model = {
    sheet: sheet.name,
    address: F.splitAddress(cell.address).address,
    formula: null,
    text: cell.text[0][0],
    sources: [],
    notes: [],
  };
  state.current = { sheet: model.sheet, address: model.address };
  const formula = cell.formulas[0][0];
  if (typeof formula !== "string" || !formula.startsWith("=")) return model;
  model.formula = String(cell.formulasLocal[0][0]);

  const sheetNames = new Map(wb.worksheets.items.map((w) => [w.name.toLowerCase(), w.name]));
  const realSheet = (name) => sheetNames.get(name.toLowerCase()) || name;
  const parsed = F.parse(model.formula);
  const sources = new Map();
  const add = (key, init, span) => {
    let src = sources.get(key);
    if (!src) {
      src = { key, spans: [], order: Infinity, color: NEUTRAL, ...init };
      sources.set(key, src);
    }
    if (span) {
      src.spans.push(span);
      src.order = Math.min(src.order, span[0]);
    }
    return src;
  };

  // 1. Rechtstreekse verwijzingen, in de volgorde van de formule.
  for (const ref of parsed.refs) {
    const span = [ref.start, ref.end];
    if (ref.external || ref.threeD) {
      add(`x:${ref.text}`, {
        label: ref.text,
        problem: ref.external
          ? "Verwijst naar een andere werkmap; die kan niet getoond worden."
          : "Verwijst naar meerdere bladen tegelijk (3D-verwijzing); die wordt niet getoond.",
      }, span);
      continue;
    }
    const area = F.parseArea(ref.address);
    if (!area) continue;
    const name = ref.sheet === null ? model.sheet : realSheet(ref.sheet);
    const key = sourceKey(name, area) + (ref.spill ? "#" : "");
    add(key, { sheet: name, area, spill: ref.spill, colored: name === model.sheet }, span);
  }

  // 2. Gedefinieerde namen (van het actieve blad of de werkmap).
  const names = new Map();
  for (const nm of parsed.names) {
    if (nm.sheet && nm.sheet.toLowerCase() !== model.sheet.toLowerCase()) continue;
    const lower = nm.name.toLowerCase();
    if (!names.has(lower)) {
      if (names.size >= MAX_NAMES) continue;
      const local = sheet.names.getItemOrNullObject(nm.name);
      const global = nm.sheet ? null : wb.names.getItemOrNullObject(nm.name);
      local.load("type");
      if (global) global.load("type");
      names.set(lower, { name: nm.name, local, global, spans: [] });
    }
    names.get(lower).spans.push([nm.start, nm.end]);
  }
  if (names.size) {
    try {
      await ctx.sync();
    } catch (e) {
      console.warn("Namen opzoeken mislukt", e);
      names.clear();
    }
    if (stale()) return null;
  }

  const lookups = [];
  for (const g of names.values()) {
    const item = !g.local.isNullObject ? g.local : g.global && !g.global.isNullObject ? g.global : null;
    if (!item || item.type !== "Range") continue;
    const range = item.getRangeOrNullObject();
    range.load("address");
    lookups.push({ g, range });
  }
  const spills = [];
  if (apiSupported("1.12")) {
    for (const src of sources.values()) {
      if (!src.spill || !sheetNames.has(src.sheet.toLowerCase())) continue;
      const range = wb.worksheets.getItem(src.sheet)
        .getRangeByIndexes(src.area.r1, src.area.c1, 1, 1)
        .getSpillingToRangeOrNullObject();
      range.load("address");
      spills.push({ src, range });
    }
  }
  if (lookups.length || spills.length) {
    await ctx.sync();
    if (stale()) return null;
  }
  for (const { g, range } of lookups) {
    if (range.isNullObject) continue;
    const { sheet: s, address } = F.splitAddress(range.address);
    const area = F.parseArea(address);
    if (!area || !s) continue;
    const src = add(`naam:${g.name.toLowerCase()}`, { sheet: realSheet(s), area, via: `naam ${g.name}` }, null);
    for (const span of g.spans) add(src.key, {}, span);
  }
  for (const { src, range } of spills) {
    if (range.isNullObject) continue;
    const area = F.parseArea(F.splitAddress(range.address).address);
    if (area) src.area = area;
  }

  // 3. Tabelverwijzingen (Tabel1[Kolom]) laat Excel zelf oplossen.
  if (parsed.hasStructured && apiSupported("1.12")) {
    const precedents = cell.getDirectPrecedents();
    precedents.load("addresses");
    try {
      await ctx.sync();
      for (const list of precedents.addresses) {
        for (const { sheet: s, address } of F.splitAddressList(list)) {
          const area = F.parseArea(address);
          if (!area || !s) continue;
          const name = realSheet(s);
          const known = [...sources.values()].some((x) => x.area && x.sheet === name && F.contains(x.area, area));
          if (!known) add(sourceKey(name, area), { sheet: name, area, via: "tabel" }, null);
        }
      }
    } catch (e) {
      console.warn("Tabelverwijzingen opzoeken mislukt", e);
      model.notes.push("De tabelverwijzingen in deze formule konden niet opgezocht worden.");
    }
    if (stale()) return null;
  }

  // Volgorde en kleuren zoals in de formulebalk.
  const list = [...sources.values()].sort((a, b) => a.order - b.order);
  const colors = F.assignColors(list.map((src) => (src.colored ? src.key : null)), REF_COLORS.length);
  for (const src of list) {
    if (colors.has(src.key)) src.color = REF_COLORS[colors.get(src.key)];
    if (src.area && !sheetNames.has(src.sheet.toLowerCase())) src.problem = `Het blad "${src.sheet}" bestaat niet.`;
  }
  model.sources = list;
  if (list.length > MAX_CARDS) {
    model.notes.push(`Alleen de eerste ${MAX_CARDS} van ${list.length} verwijzingen worden getoond.`);
  }

  // 4. Uitsnede per broncel: tekst, kolombreedtes, rijhoogtes en de broncel zelf.
  const shown = list.slice(0, MAX_CARDS).filter((src) => src.area && !src.problem);
  const proxies = new Map();
  for (const src of shown) {
    const ws = wb.worksheets.getItem(src.sheet);
    const a = src.area;
    const r2 = Math.min(a.r2, a.r1 + MAX_AREA_ROWS - 1);
    const c2 = Math.min(a.c2, a.c1 + MAX_AREA_COLS - 1);
    const cut = { rows: r2 < a.r2, cols: c2 < a.c2 };
    const win = {
      r1: Math.max(0, a.r1 - state.rows),
      c1: Math.max(0, a.c1 - state.cols),
      r2: cut.rows ? r2 : Math.min(F.MAX_ROW - 1, r2 + state.rows),
      c2: cut.cols ? c2 : Math.min(F.MAX_COL - 1, c2 + state.cols),
    };
    const range = ws.getRangeByIndexes(win.r1, win.c1, win.r2 - win.r1 + 1, win.c2 - win.c1 + 1);
    range.load("text, valueTypes");
    const cols = [];
    for (let c = win.c1; c <= win.c2; c++) {
      const col = ws.getRangeByIndexes(win.r1, c, 1, 1);
      col.load("columnHidden, format/columnWidth");
      cols.push(col);
    }
    const rows = [];
    for (let r = win.r1; r <= win.r2; r++) {
      const row = ws.getRangeByIndexes(r, win.c1, 1, 1);
      row.load("rowHidden, format/rowHeight");
      rows.push(row);
    }
    const first = ws.getRangeByIndexes(a.r1, a.c1, 1, 1);
    first.load("formulasLocal, text");
    proxies.set(src, { win, cut, range, cols, rows, first });
  }
  if (shown.length) {
    await ctx.sync();
    if (stale()) return null;
  }

  // 5. De echte opmaak: Excel rendert elke uitsnede als afbeelding. Lukt dat niet, dan een eenvoudig rooster.
  let useGrid = state.mode === "grid";
  if (!useGrid && shown.length) {
    for (const p of proxies.values()) p.image = p.range.getImage();
    try {
      await ctx.sync();
    } catch (e) {
      console.warn("Afbeelding ophalen mislukt", e);
      model.notes.push("Excel kon de opmaak niet als afbeelding leveren; de snelle weergave wordt gebruikt.");
      useGrid = true;
    }
    if (stale()) return null;
  }
  if (useGrid && shown.length) {
    for (const p of proxies.values()) {
      p.props = p.range.getCellProperties({
        format: { fill: { color: true }, font: { bold: true, italic: true, color: true }, horizontalAlignment: true },
      });
    }
    try {
      await ctx.sync();
    } catch (e) {
      console.warn("Celopmaak ophalen mislukt", e);
      for (const p of proxies.values()) p.props = null;
    }
    if (stale()) return null;
  }

  for (const [src, p] of proxies) {
    src.view = {
      win: p.win,
      cut: p.cut,
      cols: p.cols.map((c, i) => ({ index: p.win.c1 + i, size: c.columnHidden ? 0 : c.format.columnWidth })),
      rows: p.rows.map((r, i) => ({ index: p.win.r1 + i, size: r.rowHidden ? 0 : r.format.rowHeight })),
      text: p.range.text,
      types: p.range.valueTypes,
      image: useGrid ? null : p.image.value,
      props: useGrid && p.props ? p.props.value : null,
    };
    src.first = { formula: String(p.first.formulasLocal[0][0]), text: String(p.first.text[0][0]) };
  }
  return model;
}

/* ---------- Weergave ---------- */

function render(model) {
  setHot(null);
  $("cell").textContent = sheetLabel(model.sheet, model.address);
  $("cell").title = $("cell").textContent;
  renderFormula(model);
  setStatus(model.notes.join(" "), null);

  const cards = $("cards");
  cards.replaceChildren();
  if (model.formula === null) {
    cards.append(h("div", { class: "empty" },
      h("strong", null, "Geen formule"),
      "Selecteer een cel met een formule om haar broncellen te zien."));
  } else if (!model.sources.length) {
    cards.append(h("div", { class: "empty" },
      h("strong", null, "Geen broncellen"),
      "Deze formule verwijst niet naar andere cellen."));
  }
  for (const src of model.sources.slice(0, MAX_CARDS)) cards.append(renderCard(src));
  updateBack();
}

function renderFormula(model) {
  const box = $("formula");
  box.replaceChildren();
  box.classList.toggle("is-value", model.formula === null);
  if (model.formula === null) {
    box.append(model.text === "" ? h("span", { class: "muted" }, "(lege cel)") : model.text);
    return;
  }
  const f = model.formula;
  const spans = [];
  for (const src of model.sources) for (const [start, end] of src.spans) spans.push({ start, end, src });
  spans.sort((a, b) => a.start - b.start);
  let pos = 0;
  for (const { start, end, src } of spans) {
    if (start < pos) continue;
    if (start > pos) box.append(breakable(f.slice(pos, start)));
    const cls = src.color === NEUTRAL ? "ref plain" : "ref";
    box.append(h("span", { class: cls, "data-key": src.key, style: { "--c": src.color } }, f.slice(start, end)));
    pos = end;
  }
  box.append(breakable(f.slice(pos)));
}

function renderCard(src) {
  const address = src.area ? F.areaToA1(src.area) : null;
  const label = src.area ? sheetLabel(src.sheet, address) : src.label;
  const tag = src.via || (src.spill ? "overloopbereik" : null);
  const card = h("article", { class: "card", "data-key": src.key, style: { "--c": src.color } },
    h("header", { class: "card-head" },
      h("span", { class: "card-addr" }, label),
      tag ? h("span", { class: "tag" }, tag) : null,
      src.area && !src.problem
        ? h("button", {
          class: "goto",
          type: "button",
          title: `Ga naar ${label}`,
          html: `${ICON_GOTO}<span>Ga naar</span>`,
          onclick: () => goTo(src.sheet, address).catch(showError),
        })
        : null));
  if (src.problem) {
    card.append(h("p", { class: "card-problem" }, src.problem));
    return card;
  }
  if (!src.view) return card;
  card.append(renderSheet(src), renderContent(src));
  const cut = [];
  if (src.view.cut.rows) cut.push(`de eerste ${MAX_AREA_ROWS} rijen`);
  if (src.view.cut.cols) cut.push(`de eerste ${MAX_AREA_COLS} kolommen`);
  if (cut.length) card.append(h("p", { class: "card-note" }, `Groot bereik: toont ${cut.join(" en ")}.`));
  return card;
}

function renderSheet(src) {
  const v = src.view;
  const a = src.area;
  const W = sum(v.cols.map((c) => c.size)) || 1;
  const H = sum(v.rows.map((r) => r.size)) || 1;
  const sizeBefore = (list, from) => sum(list.filter((x) => x.index < from).map((x) => x.size));
  const sizeWithin = (list, lo, hi) => sum(list.filter((x) => x.index >= lo && x.index <= hi).map((x) => x.size));
  const head = (x, on, label) => h("span", { class: on ? "on" : null, style: { flexGrow: x.size, display: x.size ? null : "none" } }, label);
  const rowHeadWidth = String(v.win.r2 + 1).length * 6 + 10;

  const body = h("div", { class: "sheet-body", style: { aspectRatio: `${W} / ${H}` } });
  if (v.image) body.append(h("img", { src: `data:image/png;base64,${v.image}`, alt: "" }));
  else body.append(renderGrid(v, W));
  body.append(h("div", {
    class: "mark",
    style: {
      left: pct(sizeBefore(v.cols, a.c1), W),
      top: pct(sizeBefore(v.rows, a.r1), H),
      width: pct(sizeWithin(v.cols, a.c1, a.c2), W),
      height: pct(sizeWithin(v.rows, a.r1, a.r2), H),
    },
  }));

  return h("div", {
    class: "sheet",
    style: { "--rh": `${rowHeadWidth}px`, maxWidth: `${Math.round(W * PT_TO_PX) + rowHeadWidth + 2}px` },
  },
  h("div", { class: "corner" }),
  h("div", { class: "col-heads" }, v.cols.map((c) => head(c, c.index >= a.c1 && c.index <= a.c2, F.indexToCol(c.index)))),
  h("div", { class: "row-heads" }, h("div", null, v.rows.map((r) => head(r, r.index >= a.r1 && r.index <= a.r2, r.index + 1)))),
  body);
}

// Snelle weergave: een rooster met de tekst en de belangrijkste opmaak van elke cel.
function renderGrid(v, W) {
  const grid = h("div", {
    class: "grid",
    style: {
      gridTemplateColumns: v.cols.map((c) => `${c.size}fr`).join(" "),
      gridTemplateRows: v.rows.map((r) => `${r.size}fr`).join(" "),
      fontSize: `${(11 / W) * 100}cqw`,
    },
  });
  v.rows.forEach((r, i) => v.cols.forEach((c, j) => {
    const f = v.props && v.props[i] && v.props[i][j] && v.props[i][j].format;
    const align = f && f.horizontalAlignment;
    const right = align === "Right" || (align !== "Left" && align !== "Center" && v.types[i][j] === "Double");
    const style = {};
    if (f && f.fill && f.fill.color && f.fill.color.toUpperCase() !== "#FFFFFF") style.background = f.fill.color;
    if (f && f.font) {
      if (f.font.color) style.color = f.font.color;
      if (f.font.bold) style.fontWeight = "700";
      if (f.font.italic) style.fontStyle = "italic";
    }
    if (align === "Center") style.justifyContent = "center";
    grid.append(h("div", { class: `gc${right ? " num" : ""}`, style }, v.text[i][j]));
  }));
  return grid;
}

function renderContent(src) {
  const { formula, text } = src.first;
  const isFormula = formula.startsWith("=");
  const cells = F.areaCells(src.area);
  const box = h("div", { class: "card-content" });
  if (cells > 1) {
    const first = F.areaToA1({ r1: src.area.r1, c1: src.area.c1, r2: src.area.r1, c2: src.area.c1 });
    box.append(h("div", { class: "cc-meta" }, `Bereik van ${cells.toLocaleString("nl-BE")} cellen · eerste cel ${first}`));
  }
  const shown = isFormula ? formula : text;
  box.append(h("div", { class: "cc-line" },
    h("span", { class: "fx" }, "fx"),
    shown === "" ? h("span", { class: "muted" }, "(leeg)") : h("code", { class: "src-formula" }, breakable(shown))));
  if (isFormula && text !== "") box.append(h("div", { class: "cc-result" }, `= ${text}`));
  return box;
}

/* ---------- Office ---------- */

function applyOfficeTheme() {
  try {
    const bg = Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor;
    if (!bg) return;
    const n = parseInt(bg.replace("#", ""), 16);
    const luminance = (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
    document.documentElement.dataset.theme = luminance < 0.5 ? "dark" : "light";
  } catch {
    // Thema niet beschikbaar: licht blijft.
  }
}

async function init() {
  applyOfficeTheme();
  const onSelection = async () => { if (!state.paused) scheduleRefresh(120); };
  const onChanged = async () => { if (!state.paused) scheduleRefresh(400); };
  await Excel.run(async (ctx) => {
    ctx.workbook.onSelectionChanged.add(onSelection);
    ctx.workbook.worksheets.onActivated.add(onSelection);
    if (apiSupported("1.9")) ctx.workbook.worksheets.onChanged.add(onChanged);
    await ctx.sync();
  });
  scheduleRefresh(0);
}

loadSettings();
bindUi();
updateBack();

if (window.Office) {
  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) {
      setStatus("Deze invoegtoepassing werkt enkel in Excel.", "error");
      return;
    }
    init().catch(showError);
  });
}

// Voor de ontwikkelpagina (dev/preview.html).
window.BronPanel = { render, state };
