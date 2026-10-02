const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/dcf963bb-9c61-46f5-bef2-39592a3aa903_876e3daa-0869-4ab9-88ef-368b44a1e9ef/pptx/scripts/apply_theme.js");

const OUT = path.join(__dirname, "..", "Prezentare_MPS.pptx");
const ASSETS = path.join(__dirname, "assets");

const THEME = {
  name: "MPS Blueprint",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "12263F",
    lt1: "FFFFFF",
    dk2: "46586F",
    lt2: "EEF2F7",
    accent1: "F2A900",
    accent2: "1B998B",
    accent3: "E4572E",
    accent4: "3A6EA5",
    accent5: "7A5C99",
    accent6: "9AA7B8",
    hlink: "3A6EA5",
    folHlink: "7A5C99",
  },
};
const H = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Managementul Proiectelor Software - Introducere, Proiecte si Stakeholderi";
pres.author = "Echipa MPS";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const SH = pres.ShapeType;

const W = 10;
const MX = 0.5;

pres.defineSlideMaster({
  title: "LIGHT",
  background: { color: H.lt1 },
  margin: [0.3, 0.5, 0.4, 0.5],
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 7.5, h: 0.7, fontSize: 22, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 },
        text: "Titlu",
      },
    },
    {
      text: {
        text: "MPS  ·  Introducere, Proiecte și Stakeholderi",
        options: { x: 0.5, y: 5.28, w: 5, h: 0.25, fontSize: 9, color: C.accent6, margin: 0 },
      },
    },
  ],
  slideNumber: { x: 9.0, y: 5.28, w: 0.5, h: 0.25, fontSize: 9, color: C.accent6, align: "right" },
});

pres.defineSlideMaster({
  title: "DARK",
  background: { color: H.dk1 },
  margin: [0.3, 0.5, 0.4, 0.5],
  objects: [
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.6, y: 1.2, w: 8.8, h: 1.5, fontSize: 38, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 },
        text: "Titlu",
      },
    },
    {
      text: {
        text: "MPS  ·  Introducere, Proiecte și Stakeholderi",
        options: { x: 0.5, y: 5.28, w: 5, h: 0.25, fontSize: 9, color: C.accent6, margin: 0 },
      },
    },
  ],
});

const iconCache = {};
async function icon(name, hex, size = 256) {
  const key = name + hex;
  if (iconCache[key]) return iconCache[key];
  const Comp = fa[name];
  if (!Comp) throw new Error("Missing icon " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  iconCache[key] = "image/png;base64," + buf.toString("base64");
  return iconCache[key];
}

const WHO = {
  A: { label: "A", fill: C.accent4, text: C.background1 },
  B: { label: "B", fill: C.accent2, text: C.background1 },
  AB: { label: "A + B", fill: C.accent1, text: C.text1 },
};

function chip(s, who, minutes) {
  const w = WHO[who];
  s.addShape(SH.roundRect, {
    x: 8.3, y: 0.42, w: 1.2, h: 0.32, rectRadius: 0.16,
    fill: { color: w.fill }, line: { color: w.fill, width: 0 },
    objectName: "Chip prezentator",
  });
  s.addText(`${w.label}  ·  ${minutes}`, {
    x: 8.3, y: 0.42, w: 1.2, h: 0.32, fontSize: 11, bold: true, color: w.text,
    align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Text prezentator",
  });
}

function lightSlide(title, who, minutes, section, notes) {
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  chip(s, who, minutes);
  if (notes) s.addNotes(notes);
  return s;
}

function card(s, x, y, w, h, fill, name) {
  s.addShape(SH.roundRect, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: fill || C.background2 },
    line: { color: fill || C.background2, width: 0 },
    objectName: name || "Card",
  });
}

function circleIcon(s, img, cx, cy, d, fill, name) {
  s.addShape(SH.ellipse, { x: cx, y: cy, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: (name || "Icon") + " cerc" });
  const p = d * 0.26;
  s.addImage({ data: img, x: cx + p, y: cy + p, w: d - 2 * p, h: d - 2 * p, objectName: name || "Icon" });
}

function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, valign: "top" }, o));
}

function promptBar(s, label, text, y) {
  s.addShape(SH.roundRect, { x: MX, y, w: 9, h: 0.55, rectRadius: 0.1, fill: { color: C.text1 }, line: { color: C.text1, width: 0 }, objectName: "Bara intrebare" });
  s.addText(label, {
    x: MX + 0.2, y, w: 1.35, h: 0.55, fontSize: 11, bold: true, color: C.accent1, valign: "middle", margin: 0, isTextBox: true, objectName: "Eticheta intrebare",
  });
  s.addText(text, {
    x: MX + 1.6, y, w: 7.2, h: 0.55, fontSize: 13, color: C.background1, valign: "middle", margin: 0, isTextBox: true, objectName: "Text intrebare",
  });
}

function chartFonts() {
  return {
    catAxisLabelFontFace: "+mn-lt",
    valAxisLabelFontFace: "+mn-lt",
    dataLabelFontFace: "+mn-lt",
    titleFontFace: "+mn-lt",
    legendFontFace: "+mn-lt",
  };
}

async function build() {
  const ic = {};
  const need = {
    cloud: ["FaCloud", H.lt1], sliders: ["FaSlidersH", H.lt1], finger: ["FaFingerprint", H.lt1],
    diagram: ["FaProjectDiagram", H.lt1], chart: ["FaChartLine", H.lt1], heart: ["FaHeartbeat", H.lt1],
    fileSig: ["FaFileSignature", H.lt1], hourglass: ["FaHourglassHalf", H.lt1], users: ["FaUsers", H.lt1],
    calendar: ["FaCalendarCheck", H.lt1], bullseye: ["FaBullseye", H.lt1], sync: ["FaSyncAlt", H.lt1],
    balance: ["FaBalanceScale", H.lt1], code: ["FaCode", H.lt1], cogs: ["FaCogs", H.lt1],
    plug: ["FaPlug", H.lt1], chalk: ["FaChalkboardTeacher", H.lt1], comments: ["FaComments", H.lt1],
    tie: ["FaUserTie", H.lt1], rocket: ["FaRocket", H.lt1], shield: ["FaShieldAlt", H.lt1],
    money: ["FaCoins", H.lt1], search: ["FaSearch", H.lt1], bridge: ["FaRoad", H.lt1],
    trash: ["FaTrashAlt", H.lt1], table: ["FaTable", H.lt1], mobile: ["FaMobileAlt", H.lt1],
    office: ["FaUtensils", H.lt1], tools: ["FaTools", H.lt1], vote: ["FaHandPaper", H.lt1], team: ["FaUserFriends", H.lt1],
    lightbulb: ["FaLightbulb", H.lt1], check: ["FaCheck", H.lt1], eye: ["FaEye", H.lt1],
    flag: ["FaFlagCheckered", H.lt1], dark: ["FaCommentDots", H.dk1],
  };
  for (const [k, [n, hex]] of Object.entries(need)) ic[k] = await icon(n, hex);
  const icDark = {
    users: await icon("FaUsers", H.dk1), code: await icon("FaCode", H.dk1), vote: await icon("FaHandPaper", H.dk1),
  };

  pres.addSection({ title: "Deschidere" });

  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Deschidere" });
    s.addShape(SH.ellipse, { x: 7.1, y: -1.2, w: 4.6, h: 4.6, fill: { color: C.text2, transparency: 55 }, line: { color: C.text2, width: 0 }, objectName: "Decor cerc mare" });
    s.addShape(SH.ellipse, { x: 9.0, y: 3.9, w: 2.2, h: 2.2, fill: { color: C.accent1, transparency: 70 }, line: { color: C.accent1, width: 0 }, objectName: "Decor cerc mic" });
    txt(s, "MANAGEMENTUL PROIECTELOR SOFTWARE", { x: 0.6, y: 0.75, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 3 });
    s.addText("Cod bun, proiect eșuat?", { placeholder: "title" });
    txt(s, "De la Gantt și drumul critic la stakeholderi: ce face un proiect software să reușească", {
      x: 0.6, y: 2.6, w: 6.3, h: 0.9, fontSize: 18, color: C.background2,
    });
    card(s, 0.6, 3.9, 2.4, 0.8, C.accent4, "Card prezentator A");
    txt(s, [{ text: "Prezentator A", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "Context · Istorie · CPM/PERT", options: { fontSize: 11 } }], {
      x: 0.75, y: 3.9, w: 2.2, h: 0.8, color: C.background1, valign: "middle",
    });
    card(s, 3.2, 3.9, 2.4, 0.8, C.accent2, "Card prezentator B");
    txt(s, [{ text: "Prezentator B", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "Proiecte · Stakeholderi · PM", options: { fontSize: 11 } }], {
      x: 3.35, y: 3.9, w: 2.2, h: 0.8, color: C.background1, valign: "middle",
    });
    txt(s, "25 minute · interactiv · cu dezbatere", { x: 6.0, y: 3.9, w: 3.4, h: 0.8, fontSize: 13, color: C.accent1, bold: true, valign: "middle" });
    s.addNotes(
      "A (30s). Salut. Deschidem cu o provocare: poți scrie cod impecabil și totuși să pierzi proiectul. Pe parcursul a 25 de minute vom avea 3 voturi rapide, 2 mini-exerciții și o dezbatere în două tabere. Rugăm sala să participe: nu vorbim doar noi.\n" +
        "Sursa conținutului: cursurile G-C01.00 (Introducere) și G-C01.01 (Proiecte și Stakeholderi)."
    );
  }

  {
    const s = lightSlide("Cum folosim cele 25 de minute", "AB", "1'", "Deschidere",
      "A (1 min). Arătăm harta: 5 blocuri, schimbăm prezentatorul la jumătate. Explicăm regulile jocului: la voturi ridicăm mâna, la exerciții lucrăm 2 minute în perechi, la dezbatere împărțim sala în două tabere (stânga/dreapta). Notă: dacă timpul se strânge, tăiem exercițiul de pe slide 16 și păstrăm matricea.");
    const blocks = [
      { l: "Hook", m: 3, w: "A" },
      { l: "Istorie: de la Gantt la PERT", m: 9, w: "A" },
      { l: "Proiecte", m: 5, w: "B" },
      { l: "Stakeholderi", m: 5, w: "B" },
      { l: "Dezbatere", m: 3, w: "AB" },
    ];
    let x = MX;
    const k = 9 / 25;
    blocks.forEach((b) => {
      const bw = b.m * k;
      const wc = WHO[b.w];
      s.addShape(SH.rect, { x, y: 1.5, w: bw - 0.04, h: 1.0, fill: { color: wc.fill }, line: { color: wc.fill, width: 0 }, objectName: "Bloc agenda " + b.l });
      txt(s, [{ text: b.m + "'", options: { bold: true, fontSize: 20, breakLine: true } }, { text: b.l, options: { fontSize: b.m > 4 ? 12 : 10 } }], {
        x: x + 0.08, y: 1.5, w: bw - 0.2, h: 1.0, color: wc.text, valign: "middle",
      });
      x += bw;
    });
    txt(s, "Minut 0", { x: MX, y: 2.6, w: 1, h: 0.25, fontSize: 10, color: C.text2 });
    txt(s, "Minut 25", { x: 8.5, y: 2.6, w: 1, h: 0.25, fontSize: 10, color: C.text2, align: "right" });
    const fm = [
      { i: ic.vote, t: "3 voturi rapide", d: "Mâna sus: alegeți o variantă, apoi justificăm", c: C.accent4 },
      { i: ic.lightbulb, t: "2 mini-exerciții", d: "Drum critic și identificare de stakeholderi, câte 2 minute", c: C.accent2 },
      { i: ic.comments, t: "O dezbatere", d: "Două tabere, trei teze, argumente cu exemple reale", c: C.accent3 },
    ];
    fm.forEach((f, i) => {
      const cx = MX + i * 3.05;
      card(s, cx, 3.25, 2.9, 1.7, C.background2, "Card format " + (i + 1));
      circleIcon(s, f.i, cx + 0.2, 3.45, 0.55, f.c, "Icon format " + (i + 1));
      txt(s, f.t, { x: cx + 0.2, y: 4.1, w: 2.5, h: 0.3, fontSize: 14, bold: true });
      txt(s, f.d, { x: cx + 0.2, y: 4.4, w: 2.55, h: 0.5, fontSize: 11, color: C.text2 });
    });
  }

  {
    const s = lightSlide("Cod impecabil, proiect pierdut?", "A", "2'", "Deschidere",
      "A (2 min). Idee din Motivație (2): tehnicile sunt inutile fără proces. Citim cele 3 scenarii și cerem VOT RAPID: care dintre ele ucide cele mai multe proiecte reale? Numărăm mâinile 1/2/3. Provocăm: cineva a trăit un caz? Concluzia: cerințele, testarea și livrarea sunt probleme de PROCES, nu doar de tehnică.");
    const cs = [
      { n: "1", i: ic.fileSig, c: C.accent4, t: "Cerințe scrise perfect", d: "…dar nimeni nu controlează cum se acceptă și se urmăresc modificările." },
      { n: "2", i: ic.hourglass, c: C.accent3, t: "Teste excelente", d: "…dar estimarea a fost prea optimistă, lansarea a întârziat și nu a mai rămas timp de testat." },
      { n: "3", i: ic.users, c: C.accent2, t: "Muncă fără cod", d: "Instruirea utilizatorilor, comunicarea, documentația: cineva trebuie să le planifice." },
    ];
    cs.forEach((c, i) => {
      const cx = MX + i * 3.05;
      card(s, cx, 1.3, 2.9, 2.7, C.background2, "Card scenariu " + c.n);
      circleIcon(s, c.i, cx + 0.2, 1.5, 0.7, c.c, "Icon scenariu " + c.n);
      txt(s, c.n, { x: cx + 2.2, y: 1.45, w: 0.55, h: 0.7, fontSize: 40, bold: true, color: C.accent6, align: "right" });
      txt(s, c.t, { x: cx + 0.2, y: 2.4, w: 2.5, h: 0.35, fontSize: 15, bold: true });
      txt(s, c.d, { x: cx + 0.2, y: 2.8, w: 2.5, h: 1.1, fontSize: 12, color: C.text2 });
    });
    promptBar(s, "VOT RAPID", "Care dintre cele trei strică cele mai multe proiecte reale? Mâna sus: 1, 2 sau 3.", 4.3);
  }

  {
    const s = lightSlide("De ce e special software-ul?", "A", "2'", "Deschidere",
      "A (2 min). Șase motive pentru care MPS e interesant și greu: produs intangibil, flexibil, irepetabil, proces flexibil, complexitate exponențială, sisteme critice pentru viață (control aeronave). Întrebare scurtă către sală: ce înseamnă să 'termini' un lucru intangibil? Cum arăți progresul? Facem legătura cu Gantt-ul din slide-urile următoare.");
    const items = [
      { i: ic.cloud, t: "Produs intangibil", d: "Nu-l poți vedea sau atinge; progresul e greu de măsurat", c: C.accent4 },
      { i: ic.sliders, t: "Produs flexibil", d: "Dimensiuni și constrângeri diferite de la caz la caz", c: C.accent2 },
      { i: ic.finger, t: "Adesea irepetabil", d: "Multe proiecte pot fi făcute o singură dată", c: C.accent5 },
      { i: ic.diagram, t: "Proces flexibil", d: "Nu există o singură rețetă de dezvoltare", c: C.accent4 },
      { i: ic.chart, t: "Complexitate exponențială", d: "Dimensiunea și legăturile cresc mai repede decât echipa", c: C.accent3 },
      { i: ic.heart, t: "Vieți în joc", d: "Sisteme critice: controlul unei aeronave, frânarea unui tren", c: C.accent3 },
    ];
    items.forEach((it, k) => {
      const col = k % 3;
      const row = Math.floor(k / 3);
      const cx = MX + col * 3.05;
      const cy = 1.3 + row * 1.9;
      card(s, cx, cy, 2.9, 1.75, C.background2, "Card caracteristica " + (k + 1));
      circleIcon(s, it.i, cx + 0.2, cy + 0.2, 0.6, it.c, "Icon caracteristica " + (k + 1));
      txt(s, it.t, { x: cx + 0.2, y: cy + 0.9, w: 2.5, h: 0.3, fontSize: 14, bold: true });
      txt(s, it.d, { x: cx + 0.2, y: cy + 1.2, w: 2.55, h: 0.5, fontSize: 11, color: C.text2 });
    });
  }

  pres.addSection({ title: "Istorie" });

  {
    const s = lightSlide("Scurtă istorie a managementului de proiect", "A", "2'", "Istorie",
      "A (2 min). Firul roșu: fiecare epocă a rezolvat o durere. Gantt a separat munca planificată de progres; CPM și PERT au răspuns la 'ce poate întârzia proiectul?'; anii 60 aduc standardizare (WBS, EVA), anii 70 aduc software-ul în ecuație (cascadă, Mythical Man-Month), anii 80 estimarea (Function Points, COCOMO), anii 90 calitatea totală, iar azi agilitatea și feedback-ul constant. Întrebare: care dintre aceste idei folosiți deja fără să știți numele?");
    const ev = [
      { y: "1900s", t: "Taylor", d: "Scientific Management: prima teorie despre muncă și management", c: C.accent3 },
      { y: "1910s", t: "Gantt", d: "Planificarea separă munca planificată de progres (șantiere navale)", c: C.accent4 },
      { y: "1957-58", t: "CPM și PERT", d: "Drumul critic (fabrici) și PERT (rachetele Polaris)", c: C.accent2 },
      { y: "1960s", t: "Standardizare", d: "WBS și PERT/COST obligatorii; apare EVA; PMI și IPMA", c: C.accent5 },
      { y: "1970s", t: "Intră software-ul", d: "Modelul cascadă și The Mythical Man-Month", c: C.accent3 },
      { y: "1980s", t: "Estimare", d: "Function Points și COCOMO; tool-uri PM pentru firme mici", c: C.accent4 },
      { y: "1990s", t: "Calitate totală", d: "Organizații mai flexibile, mai rapide, mai receptive", c: C.accent2 },
      { y: "Azi", t: "Agilitate", d: "Web, componente și framework-uri, feedback constant", c: C.accent1 },
    ];
    s.addShape(SH.line, { x: 0.5, y: 3.05, w: 9, h: 0, line: { color: C.accent6, width: 2 }, objectName: "Axa timp" });
    ev.forEach((e, i) => {
      const cx = 1.3 + i * 1.05;
      s.addShape(SH.ellipse, { x: cx - 0.11, y: 2.94, w: 0.22, h: 0.22, fill: { color: e.c }, line: { color: C.background1, width: 2 }, objectName: "Marcaj " + e.y });
      const above = i % 2 === 0;
      const by = above ? 1.1 : 3.4;
      txt(s, [
        { text: e.y, options: { bold: true, fontSize: 20, color: e.c, breakLine: true } },
        { text: e.t, options: { bold: true, fontSize: 13, breakLine: true } },
        { text: e.d, options: { fontSize: 11, color: C.text2 } },
      ], { x: cx - 0.85, y: by, w: 1.7, h: 1.75, valign: above ? "bottom" : "top", align: "center" });
    });
  }

  {
    const s = lightSlide("Taylor vs. Gantt: control sau încredere?", "A", "2'", "Istorie",
      "A (2 min), prima DEZBATERE scurtă. Taylor (începutul anilor 1900) pornește de la o viziune negativă: muncitorii în sarcini repetitive lucrează la minimul necesar. Contribuții reale: definirea științifică a muncii, selecția personalului, separarea responsabilităților, stimulente și pauze. Gantt (Primul Război Mondial) introduce vizualizarea planului și a progresului. ÎNTREBARE: sistemele moderne de time tracking și KPI sunt Taylor în variantă digitală? Luăm 2 opinii pro și 2 contra, apoi legăm de Gantt: vizibilitatea poate fi și instrument de colaborare, nu doar de control.");
    card(s, MX, 1.25, 4.4, 3.1, C.background2, "Card Taylor");
    s.addImage({ path: path.join(ASSETS, "taylor.png"), x: MX + 0.2, y: 1.45, w: 0.95, h: 1.72, objectName: "Foto Taylor" });
    txt(s, [{ text: "F. W. Taylor", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "Scientific Management, anii 1900", options: { fontSize: 11, color: C.text2 } }], { x: MX + 1.35, y: 1.45, w: 2.9, h: 0.7 });
    txt(s, [
      { text: "Definirea științifică a muncii", options: { bullet: true, breakLine: true } },
      { text: "Selecția științifică a personalului", options: { bullet: true, breakLine: true } },
      { text: "Stimulente și perioade de odihnă", options: { bullet: true, breakLine: true } },
      { text: "Premisă: omul lucrează la minim dacă nu e controlat", options: { bullet: true } },
    ], { x: MX + 1.35, y: 2.25, w: 2.9, h: 1.9, fontSize: 13, paraSpaceAfter: 5 });
    card(s, 5.1, 1.25, 4.4, 3.1, C.background2, "Card Gantt");
    s.addImage({ path: path.join(ASSETS, "gantt.png"), x: 5.3, y: 1.45, w: 1.2, h: 1.4, objectName: "Foto Gantt" });
    txt(s, [{ text: "Henry Gantt", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "Diagrama Gantt, Primul Război Mondial", options: { fontSize: 11, color: C.text2 } }], { x: 6.7, y: 1.45, w: 2.7, h: 0.7 });
    txt(s, [
      { text: "Vizualizează planul pe axa timpului", options: { bullet: true, breakLine: true } },
      { text: "Distinge munca planificată de progres", options: { bullet: true, breakLine: true } },
      { text: "Folosită și azi în planificare", options: { bullet: true } },
    ], { x: 6.7, y: 2.25, w: 2.7, h: 1.5, fontSize: 13, paraSpaceAfter: 5 });
    txt(s, "Instrumentul lui Gantt la lucru în șantierele navale", { x: 5.3, y: 3.75, w: 4.0, h: 0.4, fontSize: 10, color: C.text2, italic: true });
    promptBar(s, "DEZBATERE", "Time tracking-ul și KPI-urile de azi sunt Taylor în variantă digitală? Da sau nu?", 4.55);
  }

  {
    const s = lightSlide("Gantt: planul vizibil dintr-o privire", "A", "1,5'", "Istorie",
      "A (1,5 min). Exemplu de proiect software (ilustrativ). Axa orizontală este timpul, fiecare bară este o activitate, bara umplută este progresul real. Ce NU arată Gantt-ul: dependențele logice complete și ce activități chiar contează pentru termen. Exact această lipsă motivează CPM, pe care îl vedem imediat. Imaginea originală din curs (Clark, Wallace, Gantt) este în dreapta, ca să vedeți că ideea are peste un secol.");
    const tasks = [
      { n: "Cerințe", st: 0, d: 2, p: 1 },
      { n: "Design", st: 2, d: 2, p: 1 },
      { n: "Implementare", st: 4, d: 4, p: 0.6 },
      { n: "Testare", st: 7, d: 3, p: 0.1 },
      { n: "Instruire", st: 9, d: 2, p: 0 },
      { n: "Lansare", st: 11, d: 1, p: 0 },
    ];
    const gx = 2.0;
    const gy = 1.55;
    const unit = 0.3;
    const rowH = 0.5;
    for (let wk = 0; wk <= 12; wk++) {
      s.addShape(SH.line, { x: gx + wk * unit, y: gy, w: 0, h: rowH * tasks.length, line: { color: C.background2, width: 1 }, objectName: "Grila sapt " + wk });
      if (wk < 12) txt(s, String(wk + 1), { x: gx + wk * unit, y: gy - 0.28, w: unit, h: 0.22, fontSize: 9, color: C.text2, align: "center" });
    }
    txt(s, "Săptămâna", { x: MX, y: gy - 0.28, w: 1.4, h: 0.22, fontSize: 9, color: C.text2 });
    tasks.forEach((t, i) => {
      const y = gy + i * rowH;
      txt(s, t.n, { x: MX, y, w: 1.45, h: rowH, fontSize: 12, bold: true, valign: "middle" });
      s.addShape(SH.roundRect, { x: gx + t.st * unit, y: y + 0.1, w: t.d * unit, h: rowH - 0.2, rectRadius: 0.05, fill: { color: C.background2 }, line: { color: C.accent4, width: 1 }, objectName: "Bara " + t.n });
      if (t.p > 0) s.addShape(SH.roundRect, { x: gx + t.st * unit, y: y + 0.1, w: Math.max(t.d * unit * t.p, 0.1), h: rowH - 0.2, rectRadius: 0.05, fill: { color: C.accent4 }, line: { color: C.accent4, width: 0 }, objectName: "Progres " + t.n });
    });
    const todayX = gx + 6 * unit;
    s.addShape(SH.line, { x: todayX, y: gy - 0.05, w: 0, h: rowH * tasks.length + 0.1, line: { color: C.accent3, width: 2, dashType: "dash" }, objectName: "Linie azi" });
    txt(s, "Azi (săpt. 7)", { x: todayX - 0.6, y: gy + rowH * tasks.length + 0.08, w: 1.2, h: 0.25, fontSize: 10, bold: true, color: C.accent3, align: "center" });
    txt(s, "Plan vs. progres real", { x: MX, y: 4.75, w: 4, h: 0.3, fontSize: 11, color: C.text2, italic: true });
    s.addImage({ path: path.join(ASSETS, "ganttchart.png"), x: 5.9, y: 1.3, w: 3.6, h: 2.14, objectName: "Gantt original" });
    txt(s, "Sursă: The Gantt chart, a working tool of management (Clark, Wallace, Gantt)", { x: 5.9, y: 3.5, w: 3.6, h: 0.4, fontSize: 9, color: C.text2, italic: true });
    card(s, 5.9, 4.0, 3.6, 0.95, C.background2, "Card limita Gantt");
    txt(s, [{ text: "Limita: ", options: { bold: true, color: C.accent3 } }, { text: "nu spune ce activități pot întârzia întregul proiect. Aici intră CPM.", options: {} }], { x: 6.05, y: 4.0, w: 3.3, h: 0.95, fontSize: 12, valign: "middle" });
  }

  {
    const s = lightSlide("Mini-exercițiu: găsiți drumul critic", "A", "3'", "Istorie",
      "A (3 min), EXERCIȚIU INTERACTIV. Dăm 60-90 secunde sălii să adune duratele pe fiecare cale. Întrebarea din curs: 'Ce activități pot întârzia finalizarea proiectului?'. Soluție: calea A-B-D-F = 3+4+5+2 = 14 zile; calea A-C-E-F = 3+2+3+2 = 10 zile. Drumul critic este A-B-D-F (nicio zi de rezervă). C și E au slack de 4 zile. Valori ES/EF afișate pe noduri (ES = start cel mai devreme, EF = final cel mai devreme). Observații: (1) CPM (1957) e determinist, o singură estimare per activitate. (2) PERT (1958) folosește 3 estimări: (O + 4M + P) / 6. (3) Dacă 'comprimăm' B cu 2 zile, drumul critic poate să se mute pe A-C-E-F? Nu: A-B-D-F ar deveni 12, iar A-C-E-F rămâne 10, deci același drum rămâne critic. Dacă comprimăm cu 5 zile, devine 9 și se schimbă drumul. Bonus pentru discuție.");
    txt(s, "Fiecare nod: activitate · durată (zile) · ES/EF. Care este drumul cu durata cea mai mare?", { x: MX, y: 0.95, w: 9, h: 0.3, fontSize: 12, color: C.text2 });
    const nw = 1.5;
    const nh = 0.85;
    const cols = [0.5, 2.75, 5.0, 7.25];
    const nodes = {
      A: { x: cols[0], y: 2.4, d: 3, es: 0, ef: 3, crit: true },
      B: { x: cols[1], y: 1.5, d: 4, es: 3, ef: 7, crit: true },
      C: { x: cols[1], y: 3.3, d: 2, es: 3, ef: 5, crit: false },
      D: { x: cols[2], y: 1.5, d: 5, es: 7, ef: 12, crit: true },
      E: { x: cols[2], y: 3.3, d: 3, es: 5, ef: 8, crit: false },
      F: { x: cols[3], y: 2.4, d: 2, es: 12, ef: 14, crit: true },
    };
    const edges = [["A", "B"], ["A", "C"], ["B", "D"], ["C", "E"], ["D", "F"], ["E", "F"]];
    edges.forEach(([a, b]) => {
      const na = nodes[a];
      const nb = nodes[b];
      const x1 = na.x + nw;
      const y1 = na.y + nh / 2;
      const x2 = nb.x;
      const y2 = nb.y + nh / 2;
      const crit = na.crit && nb.crit;
      const opts = {
        x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1),
        line: { color: crit ? C.accent3 : C.accent6, width: crit ? 3 : 1.5, endArrowType: "triangle" },
        objectName: `Sageata ${a}-${b}`,
      };
      if (y2 < y1) opts.flipV = true;
      s.addShape(SH.line, opts);
    });
    Object.entries(nodes).forEach(([k, n]) => {
      const fill = n.crit ? C.accent3 : C.accent4;
      s.addShape(SH.roundRect, { x: n.x, y: n.y, w: nw, h: nh, rectRadius: 0.1, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: "Nod " + k });
      txt(s, [
        { text: `${k}  ·  ${n.d} zile`, options: { bold: true, fontSize: 14, breakLine: true } },
        { text: `ES ${n.es}  ·  EF ${n.ef}`, options: { fontSize: 11 } },
      ], { x: n.x, y: n.y, w: nw, h: nh, color: C.background1, align: "center", valign: "middle" });
    });
    card(s, MX, 4.4, 4.4, 0.75, C.background2, "Card rezultat critic");
    txt(s, [{ text: "Drum critic: ", options: { bold: true, color: C.accent3 } }, { text: "A → B → D → F = 14 zile, fără rezervă", options: {} }], { x: MX + 0.15, y: 4.4, w: 4.2, h: 0.75, fontSize: 13, valign: "middle" });
    card(s, 5.1, 4.4, 4.4, 0.75, C.background2, "Card rezultat slack");
    txt(s, [{ text: "Slack: ", options: { bold: true, color: C.accent4 } }, { text: "C și E pot întârzia 4 zile fără efect", options: {} }], { x: 5.25, y: 4.4, w: 4.2, h: 0.75, fontSize: 13, valign: "middle" });
  }

  {
    const s = lightSlide("Legea lui Brooks și explozia canalelor", "A", "2,5'", "Istorie",
      "A (2,5 min), încheiem blocul A. The Mythical Man-Month (Brooks, 1975): adăugarea de oameni într-un proiect software întârziat îl întârzie și mai mult. Motivul matematic: canalele de comunicare cresc ca n(n-1)/2, deci pătratic. 5 oameni = 10 canale, 20 de oameni = 190 canale. Plus costul de instruire a noilor veniți. DEZBATERE RAPIDĂ: este legea lui Brooks încă valabilă cu microservicii, documentație bună și tool-uri de colaborare? Luăm 2 argumente. Predăm apoi lui B: 'dacă proiectele sunt atât de grele de gestionat, hai să definim ce este de fapt un proiect'.");
    s.addChart(pres.charts.BAR, [{ name: "Canale de comunicare", labels: ["3", "5", "8", "10", "15", "20"], values: [3, 10, 28, 45, 105, 190] }], {
      x: MX, y: 1.2, w: 5.6, h: 3.7, barDir: "col",
      chartColors: [H.accent4],
      showTitle: true, title: "Canale de comunicare = n(n−1)/2", titleFontSize: 13, titleColor: H.dk1,
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: H.dk1, dataLabelFontSize: 11,
      catAxisLabelColor: H.dk2, valAxisLabelColor: H.dk2, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
      valGridLine: { color: H.lt2, size: 1 }, catGridLine: { style: "none" },
      showLegend: false, showCatAxisTitle: true, catAxisTitle: "Dimensiunea echipei (persoane)", catAxisTitleFontSize: 10, catAxisTitleColor: H.dk2,
      ...chartFonts(),
    });
    s.addImage({ path: path.join(ASSETS, "mmm.png"), x: 6.4, y: 1.3, w: 0.8, h: 1.17, objectName: "Coperta Mythical Man-Month" });
    txt(s, [{ text: "Mai mulți oameni într-un proiect software întârziat îl întârzie și mai mult.", options: { italic: true, fontSize: 12, breakLine: true } }, { text: "F. Brooks, The Mythical Man-Month", options: { fontSize: 10, color: C.text2 } }], { x: 7.35, y: 1.3, w: 2.15, h: 1.3 });
    card(s, 6.4, 2.85, 3.1, 2.05, C.text1, "Card dezbatere Brooks");
    txt(s, [
      { text: "DEZBATERE", options: { bold: true, fontSize: 11, color: C.accent1, breakLine: true } },
      { text: "Mai e valabilă legea lui Brooks cu microservicii, documentație bună și tool-uri moderne?", options: { fontSize: 13, color: C.background1, breakLine: true } },
      { text: "Stânga: da   ·   Dreapta: nu", options: { fontSize: 11, bold: true, color: C.accent1 } },
    ], { x: 6.6, y: 3.0, w: 2.75, h: 1.8, paraSpaceAfter: 6 });
  }

  pres.addSection({ title: "Proiecte" });

  {
    const s = lightSlide("Ce este, de fapt, un proiect?", "B", "1,5'", "Proiecte",
      "B (1,5 min), preia scena. Definiția (PMBOK): efort temporar pentru un produs, serviciu sau rezultat unic. Cele patru caracteristici: TEMPORAR (început și sfârșit; rezultatele pot dura), UNIC, ELABORARE PROGRESIVĂ (pași și incremente, nu totul definit de la început), CONSTRÂNGERI de resurse. Nuanță importantă: un proiect se poate termina și prin eșec (obiective neîndeplinite). Imaginea din curs arată ciclul de viață: inițiere, planificare, execuție + monitorizare, închidere.");
    card(s, MX, 1.2, 9, 1.05, C.text1, "Definitie proiect");
    txt(s, [{ text: "„Un efort temporar realizat pentru crearea unui produs, serviciu sau rezultat unic.”", options: { italic: true, fontSize: 17, color: C.background1 } }, { text: "   PMBOK", options: { bold: true, fontSize: 11, color: C.accent1 } }], { x: MX + 0.25, y: 1.2, w: 8.5, h: 1.05, valign: "middle" });
    const chs = [
      { i: ic.calendar, t: "Temporar", d: "Început și sfârșit; poate sfârși și prin eșec", c: C.accent4 },
      { i: ic.finger, t: "Unic", d: "Produs, serviciu sau cunoștințe livrate o singură dată", c: C.accent5 },
      { i: ic.sync, t: "Elaborare progresivă", d: "Pași și incremente, nu totul definit din start", c: C.accent2 },
      { i: ic.balance, t: "Constrângeri", d: "Timp, buget și oameni limitați", c: C.accent3 },
    ];
    chs.forEach((c, i) => {
      const cx = MX + i * 2.28;
      card(s, cx, 2.5, 2.15, 1.35, C.background2, "Card caracteristica proiect " + (i + 1));
      circleIcon(s, c.i, cx + 0.15, 2.62, 0.5, c.c, "Icon caracteristica proiect " + (i + 1));
      txt(s, c.t, { x: cx + 0.75, y: 2.62, w: 1.35, h: 0.5, fontSize: 13, bold: true, valign: "middle" });
      txt(s, c.d, { x: cx + 0.15, y: 3.2, w: 1.9, h: 0.6, fontSize: 10.5, color: C.text2 });
    });
    s.addImage({ path: path.join(ASSETS, "lifecycle.png"), x: 5.4, y: 4.0, w: 2.2, h: 1.26, objectName: "Ciclu de viata proiect" });
    txt(s, [{ text: "Întrebare: ", options: { bold: true, color: C.accent3 } }, { text: "dacă un proiect e temporar, de ce un sistem software livrat durează ani de zile?", options: {} }], { x: MX, y: 4.1, w: 4.7, h: 0.9, fontSize: 13, valign: "middle" });
    txt(s, "Ciclul de viață al proiectului", { x: 7.7, y: 4.4, w: 1.8, h: 0.5, fontSize: 10, italic: true, color: C.text2 });
  }

  {
    const s = lightSlide("Proiect sau operațional? Votați!", "B", "2'", "Proiecte",
      "B (2 min), VOT RAPID pe fiecare element (P = proiect, O = operațional). Răspunsuri sugerate: Pregătirea cinei: O dacă e rutină zilnică, P dacă e un banchet unic. Fabricarea unei mașini: O (producție în serie). Design-ul unei mașini: P. Redactarea unui articol: P. Dezvoltarea unui sistem software: P. Menținerea unui sistem software: O (dar o versiune majoră poate fi tratată ca proiect). Managementul personalului: O. Ideea: ambele sunt făcute de oameni, cu resurse limitate, planificate, executate și controlate. Diferența: proiectul urmărește un obiectiv și se încheie; operațiunile susțin business-ul. Insistați pe cazurile de graniță: acolo apare discuția.");
    const ex = [
      { t: "Pregătirea cinei", i: ic.office },
      { t: "Fabricarea unei mașini", i: ic.cogs },
      { t: "Design-ul unei mașini", i: ic.lightbulb },
      { t: "Redactarea unui articol", i: ic.fileSig },
      { t: "Dezvoltarea unui sistem software", i: ic.code },
      { t: "Menținerea unui sistem software", i: ic.tools },
      { t: "Managementul personalului", i: ic.team },
    ];
    ex.forEach((e, k) => {
      const col = k % 4;
      const row = Math.floor(k / 4);
      const cx = MX + col * 2.28;
      const cy = 1.25 + row * 1.35;
      card(s, cx, cy, 2.15, 1.2, C.background2, "Card exemplu " + (k + 1));
      circleIcon(s, e.i, cx + 0.15, cy + 0.15, 0.5, [C.accent4, C.accent2, C.accent5, C.accent3][k % 4], "Icon exemplu " + (k + 1));
      txt(s, String(k + 1), { x: cx + 1.5, y: cy + 0.1, w: 0.5, h: 0.5, fontSize: 24, bold: true, color: C.accent6, align: "right" });
      txt(s, e.t, { x: cx + 0.15, y: cy + 0.72, w: 1.9, h: 0.45, fontSize: 11.5, bold: true });
    });
    card(s, MX + 3 * 2.28, 2.6, 2.15, 1.2, C.text1, "Card legenda vot");
    txt(s, [{ text: "P = proiect", options: { bold: true, color: C.accent1, breakLine: true } }, { text: "O = operațional", options: { bold: true, color: C.background1, breakLine: true } }, { text: "Mâna sus la fiecare", options: { fontSize: 10, color: C.accent6 } }], { x: MX + 3 * 2.28 + 0.15, y: 2.6, w: 1.9, h: 1.2, fontSize: 13, valign: "middle" });
    promptBar(s, "ÎN COMUN", "Oameni, resurse limitate, planificare și control. Diferă ținta: finalizare vs. susținerea business-ului.", 4.2);
  }

  {
    const s = lightSlide("Subproiect, proiect, program, portofoliu", "B", "1'", "Proiecte",
      "B (1 min). Scară de agregare: subproiectul este o parte administrată ca proiect; proiectul produce un rezultat unic; programul grupează proiecte înrudite gestionate coordonat pentru un beneficiu; portofoliul grupează proiecte/programe fără legătură între ele, doar pentru a fi gestionate strategic. Exemplu rapid: o bancă are un portofoliu (mobile banking, migrare cloud, conformitate), iar 'mobile banking' este un program cu proiecte pentru iOS, Android și backend.");
    const lv = [
      { t: "Portofoliu", d: "Proiecte/programe fără legătură, grupate pentru obiective strategice", w: 9.0, c: C.text1, tc: C.background1 },
      { t: "Program", d: "Proiecte înrudite, gestionate coordonat pentru un beneficiu", w: 7.2, c: C.accent4, tc: C.background1 },
      { t: "Proiect", d: "Efort temporar pentru un rezultat unic", w: 5.4, c: C.accent2, tc: C.background1 },
      { t: "Subproiect", d: "Parte din proiect, administrată ca proiect", w: 3.6, c: C.accent1, tc: C.text1 },
    ];
    lv.forEach((l, i) => {
      const y = 1.3 + i * 0.95;
      const x = MX + (9 - l.w) / 2;
      s.addShape(SH.roundRect, { x, y, w: l.w, h: 0.82, rectRadius: 0.1, fill: { color: l.c }, line: { color: l.c, width: 0 }, objectName: "Nivel " + l.t });
      txt(s, [{ text: l.t, options: { bold: true, fontSize: 15, breakLine: true } }, { text: l.d, options: { fontSize: 11 } }], { x, y, w: l.w, h: 0.82, color: l.tc, align: "center", valign: "middle" });
    });
  }

  {
    const s = lightSlide("Cinci tipuri de proiecte software", "B", "1,5'", "Proiecte",
      "B (1,5 min). A: dezvoltare de aplicații (one-off, off-the-shelf, off-the-shelf customizat precum ERP). B: reingineria proceselor și sistemelor, ample, cer analiza situației existente, de obicei cu ERP. C: integrarea sistemelor, orizontală (sisteme similare) sau verticală (etape diferite ale unei proceduri). D: consultanță, expertiză adusă din exterior. E: instalare și training, sursă de venit și în open source. Întrebare: în ce categorie intră proiectul vostru de la facultate? Dar migrarea unei firme de la Excel la un ERP?");
    const types = [
      { l: "A", t: "Dezvoltare de aplicații", d: "One-off, off-the-shelf sau off-the-shelf customizat (ERP)", i: ic.code, c: C.accent4 },
      { l: "B", t: "Reingineria proceselor", d: "Schimbi modul de lucru al organizației; analiză amplă; adesea cu ERP", i: ic.sync, c: C.accent3 },
      { l: "C", t: "Integrarea sistemelor", d: "Orizontală (sisteme similare) sau verticală (etape ale unei proceduri)", i: ic.plug, c: C.accent2 },
      { l: "D", t: "Consultanță", d: "Expertiză adusă din exterior echipei", i: ic.tie, c: C.accent5 },
      { l: "E", t: "Instalare și training", d: "Sursă de venit și în open source", i: ic.chalk, c: C.accent4 },
    ];
    types.forEach((t, k) => {
      const y = 1.2 + k * 0.76;
      card(s, MX, y, 9, 0.68, C.background2, "Card tip " + t.l);
      circleIcon(s, t.i, MX + 0.12, y + 0.09, 0.5, t.c, "Icon tip " + t.l);
      txt(s, t.l, { x: MX + 0.8, y, w: 0.4, h: 0.68, fontSize: 22, bold: true, color: t.c, valign: "middle" });
      txt(s, t.t, { x: MX + 1.25, y, w: 2.6, h: 0.68, fontSize: 14, bold: true, valign: "middle" });
      txt(s, t.d, { x: MX + 3.9, y, w: 5.0, h: 0.68, fontSize: 12, color: C.text2, valign: "middle" });
    });
  }

  {
    const s = lightSlide("Cele 7 faze ale oricărui proiect", "B", "1'", "Proiecte",
      "B (1 min), moment de relaxare din curs. Cele 7 faze umoristice: entuziasm, deziluzie, confuzie, panică, căutarea vinovaților, pedepsirea nevinovaților, promovarea celor neparticipanți. Graficul este ILUSTRATIV, nu date reale. Întrebare: care fază v-a lovit cel mai tare la un proiect de echipă? Mesajul serios: lipsa unui proces și a comunicării duce la acest scenariu, exact de aici se justifică disciplina MPS.");
    s.addChart(pres.charts.LINE, [{ name: "Moral (ilustrativ)", labels: ["Entuziasm", "Deziluzie", "Confuzie", "Panică", "Căutarea vinovaților", "Pedepsirea nevinovaților", "Promovarea celor neparticipanți"], values: [9, 6, 4, 1.5, 2.5, 1, 7.5] }], {
      x: MX, y: 1.2, w: 9, h: 3.4,
      chartColors: [H.accent3], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 10,
      showTitle: true, title: "Moralul echipei de-a lungul proiectului (ilustrativ)", titleFontSize: 13, titleColor: H.dk1,
      showValue: false, catAxisLabelColor: H.dk1, catAxisLabelFontSize: 10, valAxisHidden: true,
      valAxisMinVal: 0, valAxisMaxVal: 10, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showLegend: false, ...chartFonts(),
    });
    txt(s, [{ text: "Glumă cu miez: ", options: { bold: true, color: C.accent3 } }, { text: "fiecare fază apare când procesul lipsește. Tot ce urmează în curs încearcă să o evite.", options: {} }], { x: MX, y: 4.7, w: 9, h: 0.4, fontSize: 12 });
  }

  pres.addSection({ title: "Stakeholderi" });

  {
    const s = lightSlide("Stakeholderii: cine câștigă sau pierde?", "B", "2'", "Stakeholderi",
      "B (2 min). Definiția PMBOK: individ sau organizație implicată activ sau al cărei interes poate fi afectat pozitiv sau negativ de proiect. Caracteristici: influențe și responsabilități diferite, roluri multiple, impact pozitiv sau negativ, greu de identificat; lipsa lor de implicare poate dăuna. Managerul și echipa sunt tot stakeholderi. Modelul cu inele din curs: INTERNI (echipa de proiect, echipa de management), DE MIJLOC (client/utilizator, sponsor, organizația participantă), EXTERNI (influenceri). Atenție la influencerii externi: nu sunt în proiect, dar îi pot schimba cursul.");
    const cx = 2.75;
    const cy = 3.25;
    const rings = [
      { r: 1.85, c: C.background2, t: "EXTERNI", tc: C.text2, dy: -1.7 },
      { r: 1.3, c: C.accent6, t: "DE MIJLOC", tc: C.text1, dy: -1.15 },
      { r: 0.75, c: C.accent4, t: "INTERNI", tc: C.background1, dy: -0.3 },
    ];
    rings.forEach((r) => {
      s.addShape(SH.ellipse, { x: cx - r.r, y: cy - r.r, w: r.r * 2, h: r.r * 2, fill: { color: r.c }, line: { color: C.background1, width: 2 }, objectName: "Inel " + r.t });
    });
    rings.forEach((r) => {
      txt(s, r.t, { x: cx - 0.9, y: cy + r.dy, w: 1.8, h: 0.3, fontSize: r.r > 1 ? 11 : 11, bold: true, color: r.tc, align: "center" });
    });
    txt(s, "Echipă +\nmanagement", { x: cx - 0.6, y: cy - 0.05, w: 1.2, h: 0.55, fontSize: 9.5, color: C.background1, align: "center" });
    const defs = [
      { t: "Interni", d: "Echipa de proiect și echipa de management al proiectului", c: C.accent4 },
      { t: "De mijloc", d: "Client/utilizator, sponsor (finanțare), organizația participantă", c: C.accent6 },
      { t: "Externi", d: "Influenceri: nu sunt în proiect, dar îi pot schimba cursul", c: C.accent3 },
    ];
    defs.forEach((d, i) => {
      const y = 1.3 + i * 0.85;
      s.addShape(SH.ellipse, { x: 5.2, y: y + 0.08, w: 0.3, h: 0.3, fill: { color: d.c }, line: { color: d.c, width: 0 }, objectName: "Marcaj " + d.t });
      txt(s, [{ text: d.t, options: { bold: true, fontSize: 14, breakLine: true } }, { text: d.d, options: { fontSize: 11, color: C.text2 } }], { x: 5.65, y, w: 3.85, h: 0.75 });
    });
    card(s, 5.2, 3.95, 4.3, 1.1, C.text1, "Card definitie stakeholder");
    txt(s, [{ text: "PMBOK: ", options: { bold: true, color: C.accent1 } }, { text: "oricine este implicat activ sau al cărui interes poate fi afectat, pozitiv sau negativ, de proiect.", options: { color: C.background1 } }], { x: 5.4, y: 3.95, w: 3.9, h: 1.1, fontSize: 12, valign: "middle" });
  }

  {
    const s = lightSlide("Exercițiu: identificați stakeholderii", "B", "2'", "Stakeholderi",
      "B (2 min), EXERCIȚIU. Împărțim sala în 5 grupuri mici (sau perechi). Fiecare ia un proiect și are 60-90 secunde să găsească 3 stakeholderi, dintre care cel puțin unul ascuns (nu sponsor, nu client). Exemple de răspunsuri: Pod către insulă: locuitorii insulei, pescarii și transportatorii pe feribot (pierd venit), autoritățile de mediu, firmele de construcții, asigurători. Groapă de gunoi: vecinii, primăria, autorități de mediu, ONG-uri, firma de salubritate, fermieri din zonă. Spreadsheet open source: dezvoltatori voluntari, comunitatea, companii care îl adoptă, concurenți comerciali. Aplicație de monitorizare a masei corporale: utilizatori, medici/nutriționiști, magazinele de aplicații, autoritatea GDPR, investitori. OpenOffice pe Android: Google, comunitatea OpenOffice, utilizatori, producători de telefoane. Tranziție: pentru aplicația de masă corporală avem o matrice putere-interes pe slide-ul următor.");
    const ps = [
      { t: "Un pod către o insulă", i: ic.bridge, c: C.accent4 },
      { t: "O groapă de gunoi", i: ic.trash, c: C.accent3 },
      { t: "Un spreadsheet open source", i: ic.table, c: C.accent2 },
      { t: "O aplicație web pentru monitorizarea masei corporale", i: ic.heart, c: C.accent5 },
      { t: "OpenOffice portat pe Android", i: ic.mobile, c: C.accent4 },
    ];
    ps.forEach((p, k) => {
      const col = k % 3;
      const row = Math.floor(k / 3);
      const cx = MX + col * 3.05;
      const cy = 1.25 + row * 1.5;
      card(s, cx, cy, 2.9, 1.35, C.background2, "Card proiect exercitiu " + (k + 1));
      circleIcon(s, p.i, cx + 0.15, cy + 0.15, 0.55, p.c, "Icon proiect exercitiu " + (k + 1));
      txt(s, String(k + 1), { x: cx + 2.25, y: cy + 0.1, w: 0.5, h: 0.5, fontSize: 26, bold: true, color: C.accent6, align: "right" });
      txt(s, p.t, { x: cx + 0.15, y: cy + 0.8, w: 2.65, h: 0.5, fontSize: 12, bold: true });
    });
    card(s, MX + 2 * 3.05, 1.25 + 1.5, 2.9, 1.35, C.text1, "Card reguli exercitiu");
    txt(s, [{ text: "2 minute, pe grupe", options: { bold: true, color: C.accent1, fontSize: 14, breakLine: true } }, { text: "Găsiți 3 stakeholderi, dintre care unul ascuns: nici sponsor, nici client.", options: { color: C.background1, fontSize: 12 } }], { x: MX + 2 * 3.05 + 0.15, y: 1.25 + 1.5, w: 2.6, h: 1.35, valign: "middle" });
    promptBar(s, "REȚINEȚI", "Un stakeholder ignorat nu dispare: apare târziu, ca risc sau ca schimbare de cerințe.", 4.35);
  }

  {
    const s = lightSlide("Putere vs. interes: pe cine ascultăm?", "B", "1,5'", "Stakeholderi",
      "B (1,5 min). Instrument practic derivat din ideea cursului că stakeholderii au influențe diferite și trebuie identificați și implicați. Exemplu aplicație de monitorizare a masei corporale. Cadranele: putere mare + interes mare = gestionează îndeaproape (sponsor); putere mare + interes mic = menține satisfăcut (magazinele de aplicații, autoritatea GDPR); putere mică + interes mare = informează (utilizatori, medici); putere mică + interes mic = monitorizează. Plasările sunt discutabile, deci ÎNTREBARE către sală: mutați pe cineva? De ce utilizatorii finali au 'putere mică' deși produsul e pentru ei (recenzii, abandon)?");
    const gx = 2.0;
    const gy = 1.2;
    const gw = 6.0;
    const gh = 3.5;
    const q = [
      { x: 0, y: 0, t: "Menține satisfăcut", c: C.background2 },
      { x: 1, y: 0, t: "Gestionează îndeaproape", c: C.accent1 },
      { x: 0, y: 1, t: "Monitorizează", c: C.background2 },
      { x: 1, y: 1, t: "Informează", c: C.background2 },
    ];
    q.forEach((c) => {
      s.addShape(SH.rect, { x: gx + c.x * gw / 2, y: gy + c.y * gh / 2, w: gw / 2 - 0.04, h: gh / 2 - 0.04, fill: { color: c.c, transparency: c.c === C.accent1 ? 70 : 0 }, line: { color: C.background1, width: 0 }, objectName: "Cadran " + c.t });
      txt(s, c.t, { x: gx + c.x * gw / 2 + 0.1, y: gy + c.y * gh / 2 + 0.08, w: gw / 2 - 0.3, h: 0.3, fontSize: 11, bold: true, color: C.text2 });
    });
    txt(s, "INTERES →", { x: gx, y: gy + gh + 0.05, w: gw, h: 0.25, fontSize: 10, bold: true, color: C.text2, align: "center" });
    s.addText("PUTERE →", { x: 0.2, y: gy + gh / 2 - 0.15, w: 1.5, h: 0.3, fontSize: 10, bold: true, color: C.text2, rotate: 270, align: "center", margin: 0, isTextBox: true, objectName: "Axa putere" });
    const pts = [
      { n: "Sponsor / investitor", px: 0.85, py: 0.2, c: C.accent3 },
      { n: "Magazine de aplicații", px: 0.2, py: 0.22, c: C.accent4 },
      { n: "Autoritatea GDPR", px: 0.28, py: 0.37, c: C.accent4 },
      { n: "Echipa de dezvoltare", px: 0.75, py: 0.4, c: C.accent2 },
      { n: "Utilizatori", px: 0.9, py: 0.62, c: C.accent2 },
      { n: "Medici / nutriționiști", px: 0.62, py: 0.78, c: C.accent2 },
      { n: "Concurenți", px: 0.2, py: 0.78, c: C.accent6 },
    ];
    pts.forEach((p) => {
      const bx = gx + p.px * gw;
      const by = gy + p.py * gh;
      s.addShape(SH.ellipse, { x: bx - 0.09, y: by - 0.09, w: 0.18, h: 0.18, fill: { color: p.c }, line: { color: C.background1, width: 1.5 }, objectName: "Punct " + p.n });
      txt(s, p.n, { x: bx - 1.0, y: by + 0.1, w: 2.0, h: 0.25, fontSize: 9.5, bold: true, align: "center" });
    });
  }

  {
    const s = lightSlide("Managerul de proiect: tehnic sau om?", "B", "1,5'", "Stakeholderi",
      "B (1,5 min), pregătește dezbaterea finală. Managerul de proiect răspunde de gestiunea proiectului și a așteptărilor stakeholderilor. Competențele din curs: comunicare și negociere, predispoziție la risc, orientare spre scop, leadership, gândire creativă, cunoștințe tehnice solide, corectitudine profesională, bun simț, stil. Grupăm în trei familii: oameni, mindset, tehnic. Observație: din nouă competențe doar una este strict tehnică. ÎNTREBARE: dacă ai putea alege un singur PM: un tehnician genial cu comunicare slabă sau un comunicator excelent cu tehnic mediu? Votați.");
    const g = [
      { t: "Oameni", c: C.accent4, i: ic.users, items: ["Comunicare și negociere", "Abilități de leader", "Bun simț și stil"] },
      { t: "Mindset", c: C.accent2, i: ic.bullseye, items: ["Orientare către scop", "Predispoziție către risc", "Gândire creativă", "Corectitudine profesională"] },
      { t: "Tehnic", c: C.accent3, i: ic.code, items: ["Cunoștințe tehnice solide"] },
    ];
    const gwid = [3.05, 3.05, 3.05];
    let x = MX;
    g.forEach((c, i) => {
      const w = gwid[i];
      card(s, x, 1.25, w - 0.15, 2.6, C.background2, "Card competente " + c.t);
      circleIcon(s, c.i, x + 0.15, 1.4, 0.55, c.c, "Icon competente " + c.t);
      txt(s, c.t, { x: x + 0.85, y: 1.4, w: w - 1.1, h: 0.55, fontSize: 16, bold: true, valign: "middle" });
      const lst = c.items.map((it, j) => ({ text: it, options: { bullet: true, breakLine: j < c.items.length - 1 } }));
      txt(s, lst, { x: x + 0.2, y: 2.1, w: w - 0.5, h: 1.65, fontSize: 13, paraSpaceAfter: 5 });
      x += w;
    });
    promptBar(s, "VOT RAPID", "Tehnician genial cu comunicare slabă sau comunicator excelent cu tehnic mediu? Stânga / dreapta.", 4.3);
  }

  pres.addSection({ title: "Dezbatere" });

  {
    const s = lightSlide("Dezbatere finală: alegeți o tabără", "AB", "3'", "Dezbatere",
      "A + B (3 min). Alegem o singură teză, după timpul rămas (de obicei 1 sau 2). Moderăm: A susține tabăra 'DA', B tabăra 'NU' pentru a deschide discuția; apoi cerem sălii argumente. TEZA 1: Un PM fără background tehnic poate conduce un proiect software. Argumente DA: coordonare, comunicare, gestionarea stakeholderilor. Argumente NU: nu poate evalua estimările, riscurile tehnice și credibilitatea față de echipă. TEZA 2: Proiectele software, fiind irepetabile și flexibile, nu pot fi planificate cu Gantt/CPM. DA: incertitudine mare, cerințe schimbătoare. NU: fără plan nu vezi abaterile; planul se actualizează, nu se elimină. TEZA 3: Adaugi oameni la un proiect întârziat? DA: pot fi preluate sarcini independente. NU: Brooks. Încheiem cu: răspunsul bun e 'depinde', dar trebuie să știm DE CE depinde.");
    const th = [
      { n: "1", t: "Un PM fără background tehnic poate conduce un proiect software.", a: "Comunică, negociază, ține stakeholderii aliniați", b: "Nu poate judeca estimările și riscurile tehnice", c: C.accent4 },
      { n: "2", t: "Software-ul fiind irepetabil și flexibil, planurile Gantt/CPM sunt inutile.", a: "Cerințele se schimbă, estimările mint", b: "Fără plan nu vezi abaterile; planul se actualizează", c: C.accent2 },
      { n: "3", t: "Un proiect întârziat se salvează cu mai mulți oameni.", a: "Se pot delega sarcini independente", b: "Legea lui Brooks: instruire + canale de comunicare", c: C.accent3 },
    ];
    th.forEach((t, i) => {
      const y = 1.2 + i * 1.3;
      card(s, MX, y, 9, 1.2, C.background2, "Card teza " + t.n);
      s.addShape(SH.ellipse, { x: MX + 0.15, y: y + 0.3, w: 0.6, h: 0.6, fill: { color: t.c }, line: { color: t.c, width: 0 }, objectName: "Numar teza " + t.n });
      txt(s, t.n, { x: MX + 0.15, y: y + 0.3, w: 0.6, h: 0.6, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle" });
      txt(s, t.t, { x: MX + 0.95, y: y + 0.08, w: 3.55, h: 1.05, fontSize: 13, bold: true, valign: "middle" });
      txt(s, [{ text: "DA: ", options: { bold: true, color: C.accent2 } }, { text: t.a, options: { breakLine: true } }, { text: "NU: ", options: { bold: true, color: C.accent3 } }, { text: t.b, options: {} }], { x: 5.2, y: y + 0.08, w: 4.15, h: 1.05, fontSize: 12, valign: "middle", paraSpaceAfter: 4 });
    });
  }

  {
    const s = lightSlide("Cadrul de gestiune a proiectului", "AB", "1'", "Dezbatere",
      "A + B (1 min). Rezumăm: cele 8 preocupări ale unui framework de dezvoltare software sunt chiar ce vom învăța în restul cursului: fezabilitate, scopuri, timp, costuri, controlul schimbărilor și configurații, calitate, riscuri, resurse umane. Gantt, CPM și PERT acoperă timpul; stakeholderii leagă scopul de oameni; Brooks atinge resursele umane. Cadrul trebuie să creeze viziune comună, să structureze lucrul progresiv, să reducă nesiguranța, să evidențieze abaterile, să asigure coerența și să motiveze echipa.");
    const pr = [
      { t: "Fezabilitate", i: ic.search, c: C.accent4 },
      { t: "Scopuri", i: ic.bullseye, c: C.accent2 },
      { t: "Timp", i: ic.hourglass, c: C.accent3 },
      { t: "Costuri", i: ic.money, c: C.accent5 },
      { t: "Schimbări și configurații", i: ic.sync, c: C.accent4 },
      { t: "Calitate", i: ic.shield, c: C.accent2 },
      { t: "Riscuri", i: ic.heart, c: C.accent3 },
      { t: "Resurse umane", i: ic.users, c: C.accent5 },
    ];
    pr.forEach((p, k) => {
      const col = k % 4;
      const row = Math.floor(k / 4);
      const cx = MX + col * 2.28;
      const cy = 1.25 + row * 1.35;
      card(s, cx, cy, 2.15, 1.2, C.background2, "Card preocupare " + (k + 1));
      circleIcon(s, p.i, cx + 0.15, cy + 0.15, 0.55, p.c, "Icon preocupare " + (k + 1));
      txt(s, p.t, { x: cx + 0.15, y: cy + 0.78, w: 1.9, h: 0.38, fontSize: 12.5, bold: true });
    });
    card(s, MX, 4.1, 9, 0.9, C.text1, "Card legatura cu cursul");
    txt(s, [{ text: "Gantt, CPM, PERT ", options: { bold: true, color: C.accent1 } }, { text: "răspund la timp.  ", options: { color: C.background1 } }, { text: "Stakeholderii ", options: { bold: true, color: C.accent1 } }, { text: "leagă scopul de oameni.  ", options: { color: C.background1 } }, { text: "Brooks ", options: { bold: true, color: C.accent1 } }, { text: "arată limita resurselor umane.", options: { color: C.background1 } }], { x: MX + 0.25, y: 4.1, w: 8.5, h: 0.9, fontSize: 13, valign: "middle" });
  }

  {
    const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Dezbatere" });
    s.addText("Ce luăm cu noi", { placeholder: "title", y: 0.6, h: 0.9 });
    const tk = [
      { t: "Procesul bate tehnica", d: "Cod bun fără proces controlat nu garantează un produs livrat" },
      { t: "Planul e instrument, nu dogmă", d: "De la Gantt la CPM și PERT: ce contează pentru termen" },
      { t: "Oamenii sunt proiectul", d: "Stakeholderii și echipa decid reușita mai des decât tehnologia" },
    ];
    tk.forEach((t, i) => {
      const cx = 0.6 + i * 3.0;
      s.addShape(SH.roundRect, { x: cx, y: 1.9, w: 2.85, h: 1.9, rectRadius: 0.1, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, objectName: "Card concluzie " + (i + 1) });
      txt(s, String(i + 1), { x: cx + 0.2, y: 2.0, w: 0.6, h: 0.6, fontSize: 32, bold: true, color: C.accent1 });
      txt(s, [{ text: t.t, options: { bold: true, fontSize: 14, breakLine: true } }, { text: t.d, options: { fontSize: 11, color: C.background2 } }], { x: cx + 0.2, y: 2.65, w: 2.5, h: 1.1, color: C.background1, paraSpaceAfter: 4 });
    });
    txt(s, "Întrebări?", { x: 0.6, y: 4.1, w: 5, h: 0.7, fontSize: 34, bold: true, color: C.accent1, fontFace: "Cambria" });
    txt(s, "Mulțumim pentru participare", { x: 0.6, y: 4.8, w: 5, h: 0.3, fontSize: 13, color: C.accent6 });
    s.addNotes("A + B (rezervă 0-1 min). Rezumăm cele 3 idei și deschidem pentru întrebări. Dacă timpul s-a consumat, spunem doar cele 3 idei și mulțumim.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
}

build().catch((e) => {
  console.error(e);
  process.exit(1);
});
