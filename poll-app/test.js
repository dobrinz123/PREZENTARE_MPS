import assert from "node:assert/strict";
import { handle } from "./lib/core.js";
import { createMemoryStore } from "./lib/store.js";

const store = createMemoryStore();
const env = { MASTER_PIN: "9999" };
const get = (query = {}) => handle({ method: "GET", query, headers: {}, body: null }, store, env);
const post = (body, pin) => handle({ method: "POST", query: {}, headers: pin ? { "x-master-pin": pin } : {}, body }, store, env);

let r = await get();
assert.equal(r.status, 200);
assert.equal(r.body.ctrl.status, "idle");
assert.equal(r.body.polls.length, 5);

r = await post({ action: "vote", pid: "p2", voter: "voter-0001", idx: 0 });
assert.equal(r.status, 409);

r = await post({ action: "ctrl", active: "p2", status: "open" }, "bad");
assert.equal(r.status, 401);

r = await post({ action: "ctrl", active: "p2", status: "open" }, "9999");
assert.equal(r.status, 200);

r = await post({ action: "vote", pid: "p2", voter: "voter-0001", idx: 0 });
assert.equal(r.status, 200);
r = await post({ action: "vote", pid: "p2", voter: "voter-0002", idx: 1 });
r = await post({ action: "vote", pid: "p2", voter: "voter-0002", idx: 0 });
r = await post({ action: "vote", pid: "p2", voter: "voter-0003", idx: 9 });
assert.equal(r.status, 400);
r = await post({ action: "vote", pid: "p1", voter: "voter-0003", idx: 0 });
assert.equal(r.status, 409);

r = await get({ pid: "p2" });
assert.deepEqual(r.body.counts, [2, 0]);
assert.equal(r.body.total, 2);

r = await post({ action: "ctrl", active: "p2", status: "closed" }, "9999");
r = await post({ action: "vote", pid: "p2", voter: "voter-0004", idx: 1 });
assert.equal(r.status, 409);

r = await post({ action: "clear", pid: "p2" }, "9999");
r = await get({ pid: "p2" });
assert.equal(r.body.total, 0);

r = await handle({ method: "POST", query: {}, headers: { "x-master-pin": "x" }, body: { action: "check" } }, store, {});
assert.equal(r.status, 503);

console.log("ok");
