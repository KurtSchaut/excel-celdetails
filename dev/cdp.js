// Voert JavaScript uit in het taakvenster in Excel, via de debugpoort van WebView2.
// Excel moet gestart zijn met dev/start-excel.ps1 (debugpoort 9222).
// Gebruik: node dev/cdp.js --targets
//          node dev/cdp.js "<expressie>"          (wacht op een Promise, geeft de waarde als JSON)
//          node dev/cdp.js --watch 10             (toont 10 seconden lang console-uitvoer en fouten)
const port = process.env.CDP_PORT || 9222;

async function targets() {
  return (await fetch(`http://127.0.0.1:${port}/json`)).json();
}

// Verbindt met het taakvenster; geeft { evaluate(expressie), watch(seconden), close() }.
async function connect() {
  const target = (await targets()).find((t) => t.type === "page" && t.url.includes("taskpane.html"));
  if (!target) throw new Error("Geen taakvenster gevonden. Is het paneel geopend?");
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { ws.onopen = resolve; ws.onerror = reject; });
  let nextId = 0;
  const pending = new Map();
  ws.onmessage = (event) => {
    const msg = JSON.parse(event.data);
    if (msg.id && pending.has(msg.id)) {
      pending.get(msg.id)(msg);
      pending.delete(msg.id);
    } else if (msg.method === "Runtime.consoleAPICalled") {
      const text = msg.params.args.map((a) => (a.value !== undefined ? a.value : a.description)).join(" ");
      console.log(`[${msg.params.type}] ${text}`);
    } else if (msg.method === "Runtime.exceptionThrown") {
      const d = msg.params.exceptionDetails;
      console.log(`[fout] ${d.exception ? d.exception.description : d.text}`);
    }
  };
  const send = (method, params) => new Promise((resolve) => {
    const id = ++nextId;
    pending.set(id, resolve);
    ws.send(JSON.stringify({ id, method, params }));
  });
  return {
    async evaluate(expression) {
      const reply = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      const r = reply.result || {};
      if (r.exceptionDetails) {
        const d = r.exceptionDetails;
        throw new Error(d.exception ? d.exception.description : d.text);
      }
      return r.result ? r.result.value : undefined;
    },
    async watch(seconds) {
      await send("Runtime.enable", {});
      await new Promise((r) => setTimeout(r, seconds * 1000));
    },
    close() { ws.close(); },
  };
}

async function main() {
  if (process.argv[2] === "--targets") {
    console.log(JSON.stringify((await targets()).map((t) => ({ type: t.type, title: t.title, url: t.url })), null, 2));
    return;
  }
  const pane = await connect();
  try {
    if (process.argv[2] === "--watch") await pane.watch(Number(process.argv[3]) || 10);
    else console.log(JSON.stringify(await pane.evaluate(process.argv[2]), null, 2));
  } finally {
    pane.close();
  }
}

module.exports = { connect, targets };

if (require.main === module) {
  main().catch((e) => {
    console.error(e.message || e);
    process.exit(1);
  });
}
