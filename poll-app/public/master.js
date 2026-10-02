import { el, api, post, chart } from "./common.js";

const S = { pin: "", polls: [], ctrl: { active: null, status: "idle" }, sel: "p1", counts: null, countsPid: null, storage: "" };
const view = document.getElementById("view");
const note = document.getElementById("note");

try {
  S.pin = sessionStorage.getItem("mps_pin") || "";
} catch (e) {
  S.pin = "";
}

function showNote(text, warn) {
  note.hidden = !text;
  note.className = "note" + (warn ? " warn" : "");
  note.textContent = text || "";
}

function pollById(id) {
  return S.polls.find((p) => p.id === id);
}

function renderPin() {
  const box = el("section", "card");
  box.append(el("h2", "", "Introdu PIN-ul de prezentator"));
  const row = el("form", "pin");
  const input = el("input");
  input.type = "password";
  input.id = "pin";
  input.autocomplete = "off";
  input.setAttribute("aria-label", "PIN prezentator");
  const btn = el("button", "btn", "Intră");
  btn.type = "submit";
  row.append(input, btn);
  row.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    try {
      const r = await post({ action: "check" }, input.value);
      S.pin = input.value;
      S.storage = r.storage;
      try {
        sessionStorage.setItem("mps_pin", S.pin);
      } catch (e) {
        S.pin = input.value;
      }
      showNote("");
      start();
    } catch (e) {
      showNote(e.message, true);
    }
  });
  box.append(row);
  view.replaceChildren(box);
}

async function setCtrl(active, status) {
  try {
    const r = await post({ action: "ctrl", active, status }, S.pin);
    S.ctrl = r.ctrl;
    showNote("");
  } catch (e) {
    showNote(e.message, true);
  }
  render();
  refresh();
}

async function clearVotes(pid) {
  try {
    await post({ action: "clear", pid }, S.pin);
  } catch (e) {
    showNote(e.message, true);
  }
  refresh();
}

function render() {
  const { active, status } = S.ctrl;
  const list = el("div", "list");
  S.polls.forEach((p, i) => {
    const isActive = active === p.id && status !== "idle";
    const item = el("div", "item" + (S.sel === p.id ? " sel" : ""));
    const q = el("div", "q");
    const qb = el("button");
    qb.type = "button";
    const tag = el("div", "tag", "Întrebarea " + (i + 1));
    if (isActive) tag.append(el("span", "pill " + status, status === "open" ? "deschis" : "închis"));
    qb.append(tag, el("div", "", p.q));
    qb.addEventListener("click", () => {
      S.sel = p.id;
      render();
      refresh();
    });
    q.append(qb);
    const acts = el("div", "actions");
    if (isActive && status === "open") {
      const stop = el("button", "btn stop", "Stop + arată rezultate");
      stop.type = "button";
      stop.addEventListener("click", () => setCtrl(p.id, "closed"));
      acts.append(stop);
    } else {
      const start = el("button", "btn", isActive ? "Redeschide" : "Start");
      start.type = "button";
      start.disabled = active !== null && status === "open";
      start.addEventListener("click", () => {
        S.sel = p.id;
        setCtrl(p.id, "open");
      });
      acts.append(start);
    }
    item.append(q, acts);
    list.append(item);
  });

  const sel = pollById(S.sel);
  const results = el("section", "card");
  const live = active === S.sel && status === "open";
  results.append(el("div", "tag" + (live ? " live" : ""), live ? "Rezultate live (publicul nu le vede încă)" : "Rezultate"), el("h2", "", sel.q));
  const counts = S.counts && S.countsPid === sel.id ? S.counts : sel.o.map(() => 0);
  results.append(chart(sel, counts, null));
  const row = el("div", "actions left");
  const clr = el("button", "btn alt", "Șterge voturile la această întrebare");
  clr.type = "button";
  clr.addEventListener("click", () => clearVotes(sel.id));
  const idle = el("button", "btn alt", "Ecran de așteptare pentru public");
  idle.type = "button";
  idle.addEventListener("click", () => setCtrl(null, "idle"));
  row.append(clr, idle);
  results.append(row);

  const stack = el("div", "stack");
  stack.append(list, results);
  view.replaceChildren(stack);
}

async function refresh() {
  if (!S.polls.length) return;
  try {
    const r = await api("/api/poll?pid=" + encodeURIComponent(S.sel) + "&t=" + Date.now());
    S.counts = r.counts;
    S.countsPid = r.pid;
    render();
  } catch (e) {
    showNote("Conexiune întreruptă. Încerc din nou…", true);
  }
}

async function start() {
  try {
    const data = await api("/api/poll?t=" + Date.now());
    S.polls = data.polls;
    S.ctrl = data.ctrl;
    S.storage = data.storage;
    if (S.storage === "memory" && location.hostname !== "localhost") {
      showNote("Atenție: serverul folosește memorie temporară. Voturile se pot pierde. Configurează stocarea (vezi README).", true);
    }
  } catch (e) {
    showNote("Nu mă pot conecta la server.", true);
    return;
  }
  render();
  refresh();
  setInterval(refresh, 1500);
  setInterval(async () => {
    try {
      const data = await api("/api/poll?t=" + Date.now());
      S.ctrl = data.ctrl;
    } catch (e) {
      return;
    }
  }, 3000);
}

async function boot() {
  if (!S.pin) {
    renderPin();
    return;
  }
  try {
    await post({ action: "check" }, S.pin);
    start();
  } catch (e) {
    S.pin = "";
    showNote(e.message, true);
    renderPin();
  }
}

boot();
