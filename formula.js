/* Broncellen-preview — formule-analyse.
   Zuivere functies zonder Office-afhankelijkheden, zodat ze ook in Node getest kunnen worden. */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.BronFormula = api;
})(typeof self !== "undefined" ? self : globalThis, function () {
  "use strict";

  const MAX_ROW = 1048576;
  const MAX_COL = 16384;
  const IDENT = /[\p{L}\p{N}_.$\\?]/u;
  const NAME_START = /^[\p{L}_\\]/u;
  const CELL = /^\$?([A-Z]{1,3})\$?(\d{1,7})$/i;
  const COL = /^\$?([A-Z]{1,3})$/i;
  const ROW = /^\$?(\d{1,7})$/;

  function colToIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function indexToCol(index) {
    let s = "";
    for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
      s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    }
    return s;
  }

  function cellPart(token) {
    const m = CELL.exec(token);
    if (!m) return null;
    const c = colToIndex(m[1]);
    const r = Number(m[2]) - 1;
    return c < MAX_COL && r >= 0 && r < MAX_ROW ? { r, c } : null;
  }

  function colPart(token) {
    const m = COL.exec(token);
    if (!m) return null;
    const c = colToIndex(m[1]);
    return c < MAX_COL ? c : null;
  }

  function rowPart(token) {
    const m = ROW.exec(token);
    if (!m) return null;
    const r = Number(m[1]) - 1;
    return r >= 0 && r < MAX_ROW ? r : null;
  }

  function box(r1, c1, r2, c2) {
    return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
  }

  // "B10", "$A$1:B5", "A:C", "3:5" → { r1, c1, r2, c2 } (0-gebaseerd), anders null.
  function parseArea(text) {
    const parts = text.split(":");
    if (parts.length === 1) {
      const p = cellPart(parts[0]);
      return p && box(p.r, p.c, p.r, p.c);
    }
    if (parts.length !== 2) return null;
    const [a, b] = parts;
    const p = cellPart(a), q = cellPart(b);
    if (p && q) return box(p.r, p.c, q.r, q.c);
    const ca = colPart(a), cb = colPart(b);
    if (ca !== null && cb !== null) return box(0, ca, MAX_ROW - 1, cb);
    const ra = rowPart(a), rb = rowPart(b);
    if (ra !== null && rb !== null) return box(ra, 0, rb, MAX_COL - 1);
    return null;
  }

  function areaToA1(area) {
    const { r1, c1, r2, c2 } = area;
    if (r1 === 0 && r2 === MAX_ROW - 1) return `${indexToCol(c1)}:${indexToCol(c2)}`;
    if (c1 === 0 && c2 === MAX_COL - 1) return `${r1 + 1}:${r2 + 1}`;
    const a = indexToCol(c1) + (r1 + 1);
    const b = indexToCol(c2) + (r2 + 1);
    return a === b ? a : `${a}:${b}`;
  }

  function areaCells(area) {
    return (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
  }

  function contains(outer, inner) {
    return outer.r1 <= inner.r1 && outer.c1 <= inner.c1 && outer.r2 >= inner.r2 && outer.c2 >= inner.c2;
  }

  function quoteSheet(name) {
    const plain = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) && !cellPart(name) && !/^R\d*C\d*$/i.test(name);
    return plain ? name : `'${name.replace(/'/g, "''")}'`;
  }

  // "Blad1!A1:B2" / "'Mijn blad'!C3" / "A1" → { sheet, address }
  function splitAddress(full) {
    const i = full.lastIndexOf("!");
    if (i < 0) return { sheet: null, address: full };
    let sheet = full.slice(0, i);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    return { sheet, address: full.slice(i + 1) };
  }

  // "Blad1!A1:B4, 'Mijn blad'!D1:D4" → [{ sheet, address }, …]; een gebied zonder blad erft het vorige blad.
  function splitAddressList(list) {
    const out = [];
    let current = "", quoted = false, lastSheet = null;
    const push = () => {
      const part = current.trim();
      current = "";
      if (!part) return;
      const item = splitAddress(part);
      if (item.sheet === null) item.sheet = lastSheet;
      lastSheet = item.sheet;
      out.push(item);
    };
    for (const ch of list) {
      if (ch === "'") quoted = !quoted;
      if (ch === "," && !quoted) push();
      else current += ch;
    }
    push();
    return out;
  }

  /* Zoekt de verwijzingen in een formule, in de volgorde waarin ze voorkomen.
     Werkt op zowel de Engelse als de lokale (Nederlandse) schrijfwijze: scheidingstekens en
     functienamen maken niet uit, want een naam gevolgd door "(" is altijd een functie.
     Geeft { refs, names, hasStructured } terug; start/end zijn posities in de formuletekst. */
  function parse(formula) {
    const s = String(formula);
    const n = s.length;
    const refs = [];
    const names = [];
    let hasStructured = false;

    const isIdent = (ch) => ch !== undefined && IDENT.test(ch);
    const readIdent = (j) => {
      let k = j;
      while (k < n && isIdent(s[k])) k++;
      return k;
    };

    function skipString(j) {
      for (let k = j + 1; k < n; k++) {
        if (s[k] === '"') {
          if (s[k + 1] === '"') k++;
          else return k + 1;
        }
      }
      return n;
    }

    function skipBrackets(j) {
      let depth = 0;
      for (let k = j; k < n; k++) {
        const ch = s[k];
        if (ch === "'") { k++; continue; }
        if (ch === "[") depth++;
        else if (ch === "]" && --depth === 0) return k + 1;
      }
      return n;
    }

    function skipError(j) {
      let k = j + 1;
      while (k < n && /[\p{L}\p{N}\/!?_.]/u.test(s[k])) k++;
      return k;
    }

    function addRef(start, end, sheet, address, flags) {
      let spill = false;
      if (s[end] === "#") { spill = true; end++; }
      refs.push({
        start, end, text: s.slice(start, end), sheet, address, spill,
        external: Boolean(flags && flags.external),
        threeD: Boolean(flags && flags.threeD),
      });
      return end;
    }

    // Leest een gebied ("A1", "A1:B2", "A:A", "3:5") vanaf positie j.
    function readArea(j) {
      const k = readIdent(j);
      if (s[k] === ":") {
        const k2 = readIdent(k + 1);
        const whole = s.slice(j, k2);
        if (parseArea(whole)) return { end: k2, address: whole };
      }
      const token = s.slice(j, k);
      return cellPart(token) ? { end: k, address: token } : null;
    }

    // Verwerkt wat na "Blad!" komt: een gebied, of een naam die tot dat blad behoort.
    function afterSheet(start, sheet, j, flags) {
      const area = readArea(j);
      if (area) return addRef(start, area.end, sheet, area.address, flags);
      const k = readIdent(j);
      if (flags) return addRef(start, k, sheet, s.slice(j, k), flags);
      if (k > j && NAME_START.test(s[j])) names.push({ name: s.slice(j, k), sheet, start, end: k });
      return Math.max(k, j);
    }

    let i = s.startsWith("=") ? 1 : 0;
    while (i < n) {
      const ch = s[i];

      if (ch === '"') { i = skipString(i); continue; }

      if (ch === "'") {
        let k = i + 1, name = "";
        while (k < n) {
          if (s[k] === "'") {
            if (s[k + 1] === "'") { name += "'"; k += 2; continue; }
            break;
          }
          name += s[k++];
        }
        k++;
        if (s[k] !== "!") { i = k; continue; }
        const flags = name.includes("[") ? { external: true } : name.includes(":") ? { threeD: true } : null;
        i = afterSheet(i, name, k + 1, flags);
        continue;
      }

      if (ch === "[") {
        const k = skipBrackets(i);
        const k2 = readIdent(k);
        if (s[k2] === "!") {
          i = afterSheet(i, s.slice(i, k2), k2 + 1, { external: true });
        } else {
          hasStructured = true;
          i = k;
        }
        continue;
      }

      if (ch === "#") { i = skipError(i); continue; }

      if (isIdent(ch)) {
        const k = readIdent(i);
        const token = s.slice(i, k);
        const next = s[k];
        if (next === "(") { i = k + 1; continue; }
        if (next === "[") { hasStructured = true; i = skipBrackets(k); continue; }
        if (next === "!") { i = afterSheet(i, token, k + 1, null); continue; }
        if (next === ":") {
          const k2 = readIdent(k + 1);
          if (s[k2] === "!") {
            i = afterSheet(i, s.slice(i, k2), k2 + 1, { threeD: true });
            continue;
          }
          const whole = s.slice(i, k2);
          if (parseArea(whole)) { i = addRef(i, k2, null, whole, null); continue; }
        }
        if (cellPart(token)) { i = addRef(i, k, null, token, null); continue; }
        if (NAME_START.test(token)) names.push({ name: token, sheet: null, start: i, end: k });
        i = k;
        continue;
      }

      i++;
    }

    return { refs, names, hasStructured };
  }

  /* Excel kleurt tijdens het bewerken elke verschillende verwijzing naar het actieve blad,
     in volgorde van voorkomen; dezelfde verwijzing krijgt dezelfde kleur.
     keys: de sleutels in formulevolgorde (null = niet gekleurd). Geeft Map sleutel → kleurindex. */
  function assignColors(keys, paletteSize) {
    const colors = new Map();
    for (const key of keys) {
      if (key !== null && !colors.has(key)) colors.set(key, colors.size % paletteSize);
    }
    return colors;
  }

  return {
    MAX_ROW, MAX_COL,
    colToIndex, indexToCol, parseArea, areaToA1, areaCells, contains,
    quoteSheet, splitAddress, splitAddressList, parse, assignColors,
  };
});
