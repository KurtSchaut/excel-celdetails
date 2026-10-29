// Hulpfuncties die dev/excel-test.js in het taakvenster laadt (window.__t).
window.__t = {
  sleep: (ms) => new Promise((r) => setTimeout(r, ms)),

  // Markeert de huidige weergave, zodat we kunnen zien wanneer er een nieuwe getekend is.
  mark() {
    const el = document.getElementById("cards").firstElementChild;
    if (el) el.dataset.old = "1";
  },

  async rerendered(ms = 10000) {
    const t0 = performance.now();
    while (performance.now() - t0 < ms) {
      const el = document.getElementById("cards").firstElementChild;
      if (el && !el.dataset.old && !document.body.classList.contains("busy")) return Math.round(performance.now() - t0);
      await window.__t.sleep(40);
    }
    throw new Error(`geen nieuwe weergave binnen ${ms} ms`);
  },

  async excel(fn) {
    return Excel.run(async (ctx) => {
      const out = await fn(ctx);
      await ctx.sync();
      return out;
    });
  },

  async select(sheet, address) {
    window.__t.mark();
    await window.__t.excel((ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      ws.activate();
      ws.getRange(address).select();
    });
    const ms = await window.__t.rerendered();
    return { ms, ...window.__t.snap() };
  },

  snap() {
    const $ = (id) => document.getElementById(id);
    const visible = (els) => [...els].filter((s) => s.style.display !== "none").map((s) => s.textContent);
    return {
      cell: $("cell").textContent,
      formula: $("formula").textContent,
      refs: [...document.querySelectorAll("#formula .ref")].map((r) => ({
        text: r.textContent,
        color: r.style.getPropertyValue("--c"),
        plain: r.classList.contains("plain"),
        hot: r.classList.contains("hot"),
      })),
      status: $("status").hidden ? null : $("status").textContent,
      empty: document.querySelector(".empty") ? document.querySelector(".empty").textContent : null,
      back: !$("back").disabled,
      cards: [...document.querySelectorAll(".card")].map((c) => ({
        addr: c.querySelector(".card-addr").textContent,
        color: c.style.getPropertyValue("--c"),
        tag: c.querySelector(".tag") ? c.querySelector(".tag").textContent : null,
        problem: c.querySelector(".card-problem") ? c.querySelector(".card-problem").textContent : null,
        image: Boolean(c.querySelector(".sheet-body img")),
        imageOk: Boolean(c.querySelector(".sheet-body img")) && c.querySelector(".sheet-body img").naturalWidth > 0,
        grid: Boolean(c.querySelector(".grid")),
        cols: visible(c.querySelectorAll(".col-heads span")).join(""),
        rows: visible(c.querySelectorAll(".row-heads span")).join(","),
        mark: c.querySelector(".mark") ? c.querySelector(".mark").getAttribute("style") : null,
        content: c.querySelector(".card-content") ? c.querySelector(".card-content").textContent : null,
        note: c.querySelector(".card-note") ? c.querySelector(".card-note").textContent : null,
        hot: c.classList.contains("hot"),
      })),
    };
  },
};
"ok";
