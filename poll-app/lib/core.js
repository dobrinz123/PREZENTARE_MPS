import { POLLS, findPoll } from "./polls.js";

const VOTER_RE = /^[A-Za-z0-9_-]{8,64}$/;
const STATUSES = ["idle", "open", "closed"];

function json(status, body, headers = {}) {
  return { status, body, headers };
}

function countsFor(poll, votes) {
  const counts = poll.o.map(() => 0);
  Object.values(votes).forEach((idx) => {
    if (Number.isInteger(idx) && idx >= 0 && idx < counts.length) counts[idx] += 1;
  });
  return counts;
}

function pinOk(req, env) {
  const expected = env.MASTER_PIN;
  if (!expected) return { ok: false, status: 503, error: "MASTER_PIN nu este setat pe server." };
  if (req.headers["x-master-pin"] !== expected) return { ok: false, status: 401, error: "PIN incorect." };
  return { ok: true };
}

export async function handle(req, store, env) {
  if (req.method === "GET") {
    const pid = req.query.pid;
    if (pid) {
      const poll = findPoll(pid);
      if (!poll) return json(404, { error: "Întrebare necunoscută." });
      const counts = countsFor(poll, await store.getVotes(pid));
      const total = counts.reduce((a, b) => a + b, 0);
      return json(200, { pid, counts, total }, { "cache-control": "no-store" });
    }
    const ctrl = await store.getCtrl();
    return json(
      200,
      { polls: POLLS, ctrl, storage: store.kind },
      {
        "cache-control": "public, max-age=0, s-maxage=1",
        "netlify-cdn-cache-control": "public, max-age=0, s-maxage=1",
      }
    );
  }

  if (req.method !== "POST") return json(405, { error: "Metodă nepermisă." });
  const body = req.body && typeof req.body === "object" ? req.body : {};

  if (body.action === "vote") {
    const poll = findPoll(body.pid);
    if (!poll) return json(404, { error: "Întrebare necunoscută." });
    if (!VOTER_RE.test(String(body.voter || ""))) return json(400, { error: "Identificator invalid." });
    if (!Number.isInteger(body.idx) || body.idx < 0 || body.idx >= poll.o.length) {
      return json(400, { error: "Variantă invalidă." });
    }
    const ctrl = await store.getCtrl();
    if (ctrl.active !== poll.id || ctrl.status !== "open") return json(409, { error: "Votul nu este deschis." });
    await store.castVote(poll.id, body.voter, body.idx);
    return json(200, { ok: true });
  }

  const auth = pinOk(req, env);
  if (!auth.ok) return json(auth.status, { error: auth.error });

  if (body.action === "check") return json(200, { ok: true, storage: store.kind });

  if (body.action === "ctrl") {
    const active = body.active === null ? null : String(body.active);
    if (active !== null && !findPoll(active)) return json(400, { error: "Întrebare necunoscută." });
    if (!STATUSES.includes(body.status)) return json(400, { error: "Stare invalidă." });
    const ctrl = { active: body.status === "idle" ? null : active, status: body.status, at: Date.now() };
    await store.setCtrl(ctrl);
    return json(200, { ok: true, ctrl });
  }

  if (body.action === "clear") {
    if (!findPoll(body.pid)) return json(404, { error: "Întrebare necunoscută." });
    await store.clearVotes(body.pid);
    return json(200, { ok: true });
  }

  return json(400, { error: "Acțiune necunoscută." });
}
