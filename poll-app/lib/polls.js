export const POLLS = [
  {
    id: "p1",
    q: "Care dintre cele trei strică cele mai multe proiecte reale?",
    o: [
      "Cerințe bune, dar fără control al schimbărilor",
      "Estimare prea optimistă, nu mai rămâne timp de testat",
      "Munca fără cod (instruire, comunicare) nu e planificată",
    ],
  },
  {
    id: "p2",
    q: "Time tracking-ul și KPI-urile de azi sunt Taylor în variantă digitală?",
    o: ["Da", "Nu"],
  },
  {
    id: "p3",
    q: "Mai e valabilă legea lui Brooks astăzi?",
    o: ["Da, mai mulți oameni întârzie proiectul", "Nu, tool-urile moderne o schimbă"],
  },
  {
    id: "p4",
    q: "Ce manager de proiect alegi?",
    o: ["Tehnician genial, comunicare slabă", "Comunicator excelent, tehnic mediu"],
  },
  {
    id: "p5",
    q: "La ultimul proiect de echipă care a mers prost, ce a lipsit cel mai mult?",
    o: ["Un plan clar", "Implicarea unui stakeholder cheie", "Comunicarea în echipă"],
  },
];

export function findPoll(id) {
  return POLLS.find((p) => p.id === id) || null;
}
