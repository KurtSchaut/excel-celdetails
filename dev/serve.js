// Eenvoudige statische server voor lokaal testen: node dev/serve.js [poort] [--https]
// Met --https gebruikt hij het zelfondertekende certificaat uit dev/certs (zie dev/start-excel.ps1).
const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const port = Number(process.argv[2]) || 3017;
const secure = process.argv.includes("--https");
const types = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml",
};

const handler = (req, res) => {
  console.log(new Date().toISOString().slice(11, 19), req.method, req.url, req.headers["user-agent"] || "");
  const url = decodeURIComponent(req.url.split("?")[0]);
  const file = path.join(root, url === "/" ? "dev/preview.html" : url);
  if (!file.startsWith(root)) {
    res.writeHead(403);
    res.end();
    return;
  }
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Niet gevonden");
      return;
    }
    res.writeHead(200, {
      "Content-Type": types[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
};

const server = secure
  ? https.createServer({
    key: fs.readFileSync(path.join(__dirname, "certs", "key.pem")),
    cert: fs.readFileSync(path.join(__dirname, "certs", "cert.pem")),
  }, handler)
  : http.createServer(handler);
server.on("connection", (socket) => console.log("verbinding van", socket.remoteAddress));
server.on("tlsClientError", (err) => console.log("TLS-fout:", err.code || err.message));
server.listen(port, "127.0.0.1", () => console.log(`${secure ? "https" : "http"}://localhost:${port}/`));
