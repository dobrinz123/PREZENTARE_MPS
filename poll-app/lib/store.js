const IDLE = { active: null, status: "idle", at: 0 };

function memoryStore() {
  const votes = new Map();
  let ctrl = { ...IDLE };
  return {
    kind: "memory",
    async getCtrl() {
      return ctrl;
    },
    async setCtrl(next) {
      ctrl = next;
    },
    async castVote(pid, voter, idx) {
      if (!votes.has(pid)) votes.set(pid, {});
      votes.get(pid)[voter] = idx;
    },
    async getVotes(pid) {
      return { ...(votes.get(pid) || {}) };
    },
    async clearVotes(pid) {
      votes.delete(pid);
    },
  };
}

function upstashStore(url, token) {
  async function cmd(...args) {
    const res = await fetch(url, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify(args),
    });
    if (!res.ok) throw new Error("Upstash " + res.status);
    return (await res.json()).result;
  }
  return {
    kind: "upstash",
    async getCtrl() {
      const raw = await cmd("GET", "mps:ctrl");
      return raw ? JSON.parse(raw) : { ...IDLE };
    },
    async setCtrl(next) {
      await cmd("SET", "mps:ctrl", JSON.stringify(next));
    },
    async castVote(pid, voter, idx) {
      await cmd("HSET", "mps:votes:" + pid, voter, String(idx));
    },
    async getVotes(pid) {
      const flat = (await cmd("HGETALL", "mps:votes:" + pid)) || [];
      const out = {};
      for (let i = 0; i < flat.length; i += 2) out[flat[i]] = Number(flat[i + 1]);
      return out;
    },
    async clearVotes(pid) {
      await cmd("DEL", "mps:votes:" + pid);
    },
  };
}

async function blobsStore() {
  const { getStore } = await import("@netlify/blobs");
  const store = getStore({ name: "mps-poll", consistency: "strong" });
  const prefix = (pid) => "v/" + pid + "/";
  async function keysUnder(p) {
    const { blobs } = await store.list({ prefix: p });
    return blobs.map((b) => b.key);
  }
  return {
    kind: "netlify-blobs",
    async getCtrl() {
      const raw = await store.get("ctrl", { type: "json" });
      return raw || { ...IDLE };
    },
    async setCtrl(next) {
      await store.setJSON("ctrl", next);
    },
    async castVote(pid, voter, idx) {
      const mine = await keysUnder(prefix(pid) + voter + ".");
      await Promise.all(mine.map((k) => store.delete(k)));
      await store.set(prefix(pid) + voter + "." + idx, "1");
    },
    async getVotes(pid) {
      const out = {};
      (await keysUnder(prefix(pid))).forEach((k) => {
        const rest = k.slice(prefix(pid).length);
        const dot = rest.lastIndexOf(".");
        out[rest.slice(0, dot)] = Number(rest.slice(dot + 1));
      });
      return out;
    },
    async clearVotes(pid) {
      const all = await keysUnder(prefix(pid));
      await Promise.all(all.map((k) => store.delete(k)));
    },
  };
}

let cached = null;

export async function getStoreInstance(env = process.env) {
  if (cached) return cached;
  const url = env.UPSTASH_REDIS_REST_URL || env.KV_REST_API_URL;
  const token = env.UPSTASH_REDIS_REST_TOKEN || env.KV_REST_API_TOKEN;
  if (url && token) {
    cached = upstashStore(url, token);
  } else if (env.NETLIFY || env.NETLIFY_BLOBS_CONTEXT || env.NETLIFY_LOCAL) {
    cached = await blobsStore();
  } else {
    cached = memoryStore();
  }
  return cached;
}

export function createMemoryStore() {
  return memoryStore();
}
