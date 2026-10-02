import { handle } from "../../lib/core.js";
import { getStoreInstance } from "../../lib/store.js";

export default async function poll(request) {
  const url = new URL(request.url);
  let body = null;
  if (request.method === "POST") {
    try {
      body = await request.json();
    } catch (e) {
      body = null;
    }
  }
  const req = {
    method: request.method,
    query: Object.fromEntries(url.searchParams),
    headers: Object.fromEntries([...request.headers].map(([k, v]) => [k.toLowerCase(), v])),
    body,
  };
  try {
    const out = await handle(req, await getStoreInstance(process.env), process.env);
    return new Response(JSON.stringify(out.body), {
      status: out.status,
      headers: { "content-type": "application/json; charset=utf-8", ...out.headers },
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: "Eroare de server." }), {
      status: 500,
      headers: { "content-type": "application/json; charset=utf-8" },
    });
  }
}

export const config = { path: "/api/poll" };
