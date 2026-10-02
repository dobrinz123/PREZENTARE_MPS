export function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

export async function api(path, options = {}) {
  const res = await fetch(path, options);
  let data = null;
  try {
    data = await res.json();
  } catch (e) {
    data = null;
  }
  if (!res.ok) {
    const err = new Error((data && data.error) || "Eroare de rețea.");
    err.status = res.status;
    throw err;
  }
  return data;
}

export function post(body, pin) {
  const headers = { "content-type": "application/json" };
  if (pin) headers["x-master-pin"] = pin;
  return api("/api/poll", { method: "POST", headers, body: JSON.stringify(body) });
}

export function chart(poll, counts, mineIdx) {
  const total = counts.reduce((a, b) => a + b, 0);
  const wrap = el("div", "chart");
  poll.o.forEach((label, k) => {
    const pct = total ? Math.round((counts[k] * 100) / total) : 0;
    const row = el("div", "bar-row" + (mineIdx === k ? " mine" : ""));
    const lab = el("div", "bar-label");
    lab.append(el("span", "", label + (mineIdx === k ? "  (votul tău)" : "")), el("b", "", counts[k] + " · " + pct + "%"));
    const track = el("div", "track");
    const fill = el("div", "fill");
    fill.style.width = pct + "%";
    track.append(fill);
    row.append(lab, track);
    wrap.append(row);
  });
  wrap.append(el("div", "muted", total + (total === 1 ? " vot" : " voturi")));
  return wrap;
}

export function voterId() {
  try {
    let id = localStorage.getItem("mps_voter");
    if (!id) {
      id = (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(36).slice(2)).replace(/[^A-Za-z0-9_-]/g, "");
      localStorage.setItem("mps_voter", id);
    }
    return id;
  } catch (e) {
    return "anon" + Math.random().toString(36).slice(2, 12);
  }
}

export function loadMine() {
  try {
    return JSON.parse(localStorage.getItem("mps_mine") || "{}");
  } catch (e) {
    return {};
  }
}

export function saveMine(mine) {
  try {
    localStorage.setItem("mps_mine", JSON.stringify(mine));
  } catch (e) {
    return;
  }
}
