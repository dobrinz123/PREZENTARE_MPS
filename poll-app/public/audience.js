import { el, api, post, chart, voterId, loadMine, saveMine } from "./common.js";

const voter = voterId();
const S = { polls: [], ctrl: { active: null, status: "idle" }, counts: null, countsPid: null, mine: loadMine() };
const view = document.getElementById("view");
const note = document.getElementById("note");

function showNote(text) {
  note.hidden = !text;
  note.textContent = text || "";
}

function pollById(id) {
  return S.polls.find((p) => p.id === id);
}

async function vote(pid, idx) {
  S.mine[pid] = idx;
  saveMine(S.mine);
  render();
  try {
    await post({ action: "vote", pid, voter, idx });
    showNote("");
  } catch (e) {
    showNote(e.status === 409 ? "Votul s-a închis înainte să apuci să răspunzi." : "Votul nu a putut fi trimis. Încearcă din nou.");
  }
}

function render() {
  const box = el("section", "card");
  const { active, status } = S.ctrl;
  const poll = active ? pollById(active) : null;
  if (!poll || status === "idle") {
    box.classList.add("wait");
    box.append(el("div", "tag", "Conectat"), el("h2", "", "Așteaptă următoarea întrebare"), el("div", "muted", "Prezentatorul o pornește în curând. Nu închide pagina."));
  } else if (status === "open") {
    const idx = S.polls.indexOf(poll) + 1;
    box.append(el("div", "tag live", "Întrebarea " + idx + " · votează acum"), el("h2", "", poll.q));
    poll.o.forEach((label, k) => {
      const b = el("button", "opt" + (S.mine[poll.id] === k ? " mine" : ""), label);
      b.type = "button";
      b.addEventListener("click", () => vote(poll.id, k));
      box.append(b);
    });
    box.append(el("div", "muted", "Poți schimba răspunsul până se închide votul."));
  } else {
    const idx = S.polls.indexOf(poll) + 1;
    box.append(el("div", "tag", "Întrebarea " + idx + " · rezultate"), el("h2", "", poll.q));
    if (S.counts && S.countsPid === poll.id) {
      const m = Number.isInteger(S.mine[poll.id]) ? S.mine[poll.id] : null;
      box.append(chart(poll, S.counts, m));
    } else {
      box.append(el("div", "muted", "Se încarcă rezultatele…"));
    }
  }
  view.replaceChildren(box);
}

async function tick() {
  try {
    const data = await api("/api/poll");
    S.polls = data.polls;
    S.ctrl = data.ctrl;
    if (data.ctrl.status === "closed" && data.ctrl.active) {
      const r = await api("/api/poll?pid=" + encodeURIComponent(data.ctrl.active) + "&t=" + Date.now());
      S.counts = r.counts;
      S.countsPid = r.pid;
    }
    showNote("");
  } catch (e) {
    showNote("Conexiune întreruptă. Încerc din nou…");
  }
  render();
}

render();
tick();
setInterval(tick, 2000);
