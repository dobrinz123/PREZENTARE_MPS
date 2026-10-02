import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { handle } from "./lib/core.js";
import { getStoreInstance } from "./lib/store.js";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "public");
const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8" };
const env = { MASTER_PIN: process.env.MASTER_PIN || "1234" };

async function readBody(req) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  try {
    return JSON.parse(Buffer.concat(chunks).toString() || "null");
  } catch (e) {
    return null;
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://localhost");
  if (url.pathname === "/api/poll") {
    const out = await handle(
      { method: req.method, query: Object.fromEntries(url.searchParams), headers: req.headers, body: req.method === "POST" ? await readBody(req) : null },
      await getStoreInstance({}),
      env
    );
    res.writeHead(out.status, { "content-type": "application/json", ...out.headers });
    res.end(JSON.stringify(out.body));
    return;
  }
  let file = url.pathname === "/" ? "/index.html" : url.pathname;
  if (!path.extname(file)) file += ".html";
  const full = path.join(root, path.normalize(file));
  if (!full.startsWith(root) || !fs.existsSync(full)) {
    res.writeHead(404);
    res.end("Not found");
    return;
  }
  res.writeHead(200, { "content-type": types[path.extname(full)] || "application/octet-stream" });
  fs.createReadStream(full).pipe(res);
});

const port = Number(process.env.PORT) || 3000;
server.listen(port, () => console.log("http://localhost:" + port + "  (PIN " + env.MASTER_PIN + ")"));
