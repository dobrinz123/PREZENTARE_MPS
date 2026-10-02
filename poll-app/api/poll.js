import { handle } from "../lib/core.js";
import { getStoreInstance } from "../lib/store.js";

export default async function poll(req, res) {
  const url = new URL(req.url, "http://localhost");
  const input = {
    method: req.method,
    query: Object.fromEntries(url.searchParams),
    headers: req.headers,
    body: req.body,
  };
  try {
    const out = await handle(input, await getStoreInstance(process.env), process.env);
    Object.entries(out.headers).forEach(([k, v]) => res.setHeader(k, v));
    res.status(out.status).json(out.body);
  } catch (e) {
    res.status(500).json({ error: "Eroare de server." });
  }
}
