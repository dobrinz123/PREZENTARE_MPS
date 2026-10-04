from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x12, 0x26, 0x3F)
TEAL = RGBColor(0x1B, 0x99, 0x8B)
CORAL = RGBColor(0xE4, 0x57, 0x2E)
GRAY = RGBColor(0x5B, 0x6B, 0x80)

doc = Document()
sec = doc.sections[0]
sec.left_margin = sec.right_margin = Cm(2)
sec.top_margin = sec.bottom_margin = Cm(1.8)

base = doc.styles["Normal"]
base.font.name = "Calibri"
base.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
base.font.size = Pt(11)
base.paragraph_format.space_after = Pt(3)


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = NAVY
    return p


def h2(text, tag=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(12.5)
    r.font.color.rgb = TEAL
    if tag:
        t = p.add_run("   " + tag)
        t.font.size = Pt(9.5)
        t.font.color.rgb = GRAY
    return p


def para(text, bold_lead=None, italic=False, color=None):
    p = doc.add_paragraph()
    if bold_lead:
        r = p.add_run(bold_lead + " ")
        r.bold = True
        if color:
            r.font.color.rgb = color
    r2 = p.add_run(text)
    r2.italic = italic
    return p


def bullets(items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1)
        if isinstance(it, tuple):
            r = p.add_run(it[0] + " ")
            r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)


def say(text):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = True
    c = t.rows[0].cells[0]
    shade(c, "EEF2F7")
    p = c.paragraphs[0]
    r = p.add_run("Spui: ")
    r.bold = True
    r.font.color.rgb = NAVY
    p.add_run(text).italic = True
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def qa(q, a):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run("Î: " + q)
    r.bold = True
    r.font.color.rgb = CORAL
    p2 = doc.add_paragraph()
    p2.paragraph_format.left_indent = Cm(0.5)
    p2.add_run("R: " + a)


title = doc.add_paragraph()
r = title.add_run("Ghid de învățat pentru Prezentatorul B")
r.bold = True
r.font.size = Pt(24)
r.font.color.rgb = NAVY
sub = doc.add_paragraph()
r = sub.add_run("Managementul Proiectelor Software  ·  slide-urile 10–21  ·  aproximativ 13 minute de vorbit")
r.font.size = Pt(11)
r.font.color.rgb = GRAY

h1("1. Pe scurt: ce ai de făcut")
bullets([
    ("Ce prezinți tu:", "ce este un proiect, proiect vs. operațional, tipurile de proiecte software, stakeholderii și managerul de proiect. Plus finalul împreună cu A."),
    ("Ordinea ta:", "slide 10 (proiect) → 11 (proiect sau operațional) → 12 (program, portofoliu) → 13 (tipuri) → 14 (cele 7 faze) → 15 (stakeholderi) → 16 (exercițiu) → 17 (putere–interes) → 18 (manager) → 19–21 (final cu A)."),
    ("Timp:", "10 minute pentru partea ta (cumulat: începi la 12:00, termini la 22:30), apoi 2,5 minute împreună cu A."),
    ("Regula de aur:", "nu citești slide-urile. Spui ideea cu vorbele tale, în 2–3 propoziții, apoi pui întrebarea către sală."),
    ("Dacă nu știi ceva:", "spui „Bună întrebare, în curs apare așa…” și revii la ideea principală. Nu inventa cifre sau nume."),
])

h1("2. Ideea de fond, în 30 de secunde")
para("Cursul spune așa: ca să livrezi software la timp, în buget și cu calitate, nu îți ajunge să scrii cod bun. Ai nevoie de un proces: cine face ce, până când, cu ce resurse și cine e afectat de proiect. Partea lui A arată de unde vin instrumentele (Gantt, drum critic, PERT). Partea ta arată CE este un proiect și CINE are legătură cu el (stakeholderii).")
say("Până acum am văzut cum se planifică un proiect. Eu vă arăt ce înseamnă, de fapt, un proiect și cine sunt oamenii care îl influențează sau pe care îi influențează.")

h1("3. Ce a spus A înainte (ca să nu te surprindă)")
bullets([
    ("Gantt:", "diagramă cu bare pe axa timpului. Arată planul și cât s-a făcut din el. Folosită de peste 100 de ani (Henry Gantt)."),
    ("Drum critic (CPM, 1957):", "cel mai lung lanț de activități care se succed. Dacă o activitate de pe el întârzie, întârzie tot proiectul. Activitățile din afara lui au rezervă (slack)."),
    ("PERT (1958):", "ca CPM, dar cu trei estimări per activitate (optimist, probabil, pesimist). A fost folosit la rachetele Polaris."),
    ("Legea lui Brooks:", "adăugarea de oameni într-un proiect software întârziat îl întârzie și mai mult, pentru că cresc canalele de comunicare (n(n−1)/2) și trebuie instruiți noii veniți."),
])

h1("4. Slide cu slide (partea ta)")

h2("Slide 10 · Ce este, de fapt, un proiect?", "1 min")
para("Un proiect este un efort temporar realizat pentru crearea unui produs, serviciu sau rezultat unic (definiția PMBOK).", "Ideea:")
bullets([
    ("Temporar:", "are început și sfârșit. Se poate încheia și prin eșec (obiectivele nu s-au atins). Rezultatul poate dura mult, proiectul nu."),
    ("Unic:", "produsul, serviciul sau cunoștințele livrate nu se repetă identic."),
    ("Elaborare progresivă:", "îl definești pe pași și incremente, nu totul dintr-o dată."),
    ("Constrângeri de resurse:", "timp, bani, oameni limitați."),
])
say("Un proiect e ceva cu început și sfârșit, care produce ceva unic. Dar ce rămâne după el, de exemplu o aplicație, poate fi folosit ani de zile.")
para("Echipa lansează aplicația, iar ea rămâne în folosință ani de zile. Proiectul s-a încheiat la lansare? Ce urmează după?", "Întrebarea de pe slide:")
para("Da, proiectul de dezvoltare s-a încheiat la livrare. Ce urmează (suport, corecții, adăugiri mici) este activitate operațională, adică menținerea. Dacă apar schimbări mari, se pornește un proiect nou. Slide-ul următor explică exact diferența.", "Răspuns:")

h2("Slide 11 · Proiect sau operațional?", "1,5 min")
para("Ambele sunt făcute de oameni, au resurse limitate și sunt planificate, executate și controlate. Diferența: proiectul urmărește un obiectiv și se termină; activitatea operațională susține business-ul și continuă.", "Ideea:")
bullets([
    "Pregătirea cinei: depinde. Un banchet unic este proiect, masa zilnică este operațional.",
    "Fabricarea unei mașini: operațional (producție în serie, se repetă).",
    "Design-ul unei mașini: proiect (rezultat unic).",
    "Redactarea unui articol: proiect (rezultat unic, cu termen).",
    "Dezvoltarea unui sistem software: proiect.",
    "Menținerea unui sistem software: operațional (dar o versiune majoră poate fi tratată ca proiect).",
    "Managementul personalului: operațional (activitate continuă).",
])
say("Întrebați-vă: se termină? Produce ceva unic? Dacă da la amândouă, e proiect. Dacă se repetă și susține activitatea de zi cu zi, e operațional.")
para("Pe slide verdictele sunt deja scrise. Cere sălii întâi părerea, apoi arată slide-ul. Cazurile de graniță (cina, mentenanța) sunt locul unde apare discuția, nu te teme de ele.", "Sfat:")

h2("Slide 12 · Subproiect, proiect, program, portofoliu", "0,5 min")
bullets([
    ("Subproiect:", "o parte dintr-un proiect, administrată ca proiect."),
    ("Proiect:", "efort temporar pentru un rezultat unic."),
    ("Program:", "proiecte înrudite, gestionate coordonat pentru un beneficiu comun."),
    ("Portofoliu:", "proiecte sau programe fără legătură între ele, grupate doar pentru a fi gestionate și a îndeplini obiective strategice."),
])
say("O bancă are un portofoliu cu mai multe inițiative. Aplicația mobilă este un program, iar versiunea de iOS este un proiect din el.")
para("Dacă ești în urmă cu timpul, treci repede peste acest slide. Exemplul cu banca este al nostru, nu din curs.", "Sfat:")

h2("Slide 13 · Cinci tipuri de proiecte software", "1,5 min")
bullets([
    ("A. Dezvoltare de aplicații:", "desktop, web, mobil. Pot fi one-off (făcute special pentru un client), off-the-shelf (pentru mulți utilizatori) sau off-the-shelf customizate (de exemplu sisteme ERP)."),
    ("B. Reingineria proceselor și sistemelor:", "schimbi modul în care lucrează o organizație (calitate, eficiență). Proiecte ample, cer analiza situației existente, adesea cu ERP."),
    ("C. Integrarea sistemelor:", "automatizezi fluxul de informații. Orizontală (sisteme similare) sau verticală (etape diferite ale aceleiași proceduri)."),
    ("D. Consultanță:", "aduci expertiză din exterior."),
    ("E. Instalare și training:", "servicii pentru un sistem anume; sursă de venit și în open source."),
])
para("La ce categorie intră un proiect personal pe care îl aveți sau l-ați avut?", "Întrebarea de pe slide:")
para("Ia 2–3 răspunsuri. Ghidaj: o aplicație făcută la facultate este A; trecerea unei firme de la Excel la un ERP este B; legarea a două aplicații între ele este C. Nu există un singur răspuns corect, important e să justifice.", "Cum conduci:")
para("ERP (Enterprise Resource Planning) = sistem mare care leagă procesele unei firme (contabilitate, stocuri, resurse umane) într-un singur program.", "Termen:")

h2("Slide 14 · Cele 7 faze ale oricărui proiect", "0,5 min")
para("Entuziasm → deziluzie → confuzie → panică → căutarea vinovaților → pedepsirea nevinovaților → promovarea celor neparticipanți.", "Ideea:")
say("Este o glumă din curs, dar are miez: fiecare fază apare când lipsesc procesul și comunicarea. Graficul meu este doar ilustrativ, nu date reale.")
para("Slide-ul de relaxare. Dacă ești în urmă cu timpul, îl poți sări.", "Sfat:")

h2("Slide 15 · Stakeholderii", "1,5 min")
para("Stakeholder = individ sau organizație implicată activ într-un proiect sau al cărei interes poate fi afectat, pozitiv sau negativ, de realizarea lui (PMBOK).", "Definiție:")
bullets([
    "Au influențe și responsabilități diferite și pot juca roluri diferite.",
    "Pot avea impact pozitiv sau negativ asupra proiectului.",
    "Pot fi greu de identificat, iar lipsa implicării lor poate dăuna.",
    "Managerul de proiect și echipa sunt tot stakeholderi.",
])
para("Cele trei cercuri de pe slide:", "")
bullets([
    ("Interni:", "echipa de proiect și echipa de management al proiectului."),
    ("De mijloc:", "clientul/utilizatorul, sponsorul (finanțează), organizația participantă."),
    ("Externi:", "influencerii: nu sunt în proiect, dar îi pot schimba cursul."),
])
say("Un stakeholder ignorat nu dispare. Apare târziu, ca risc sau ca schimbare de cerințe.")

h2("Slide 16 · Exercițiu: identificați stakeholderii", "1,5 min")
para("Împarte sala în grupuri mici. Fiecare ia un proiect și are ~1 minut să găsească 3 stakeholderi, dintre care unul „ascuns” (nici sponsor, nici client). Apoi ceri 2–3 răspunsuri.", "Cum conduci:")
bullets([
    ("Pod către o insulă:", "locuitorii insulei, pescarii și cei de pe feribot (pierd venit), autoritățile de mediu, firma constructoare, asiguratorii."),
    ("Groapă de gunoi:", "vecinii, primăria, autoritățile de mediu, ONG-uri, firma de salubritate, fermierii din zonă."),
    ("Spreadsheet open source:", "dezvoltatorii voluntari, comunitatea, firmele care îl adoptă, concurenții comerciali."),
    ("Aplicație web pentru masa corporală:", "utilizatorii, medici/nutriționiști, magazinele de aplicații, autoritatea de protecție a datelor, investitorii."),
    ("OpenOffice pe Android:", "Google (platforma), comunitatea OpenOffice, utilizatorii, producătorii de telefoane."),
])

h2("Slide 17 · Putere vs. interes", "1 min")
para("Atenție: acest instrument NU este din curs. L-am adăugat noi ca aplicație practică a ideii din curs că stakeholderii au influențe diferite și trebuie identificați și implicați. Spune asta dacă te întreabă cineva de unde vine.", "Important:", color=CORAL)
bullets([
    ("Putere mare + interes mare:", "gestionezi îndeaproape (sponsorul)."),
    ("Putere mare + interes mic:", "menții mulțumit (magazinele de aplicații, autoritatea de date)."),
    ("Putere mică + interes mare:", "informezi (utilizatorii, medicii)."),
    ("Putere mică + interes mic:", "monitorizezi (concurenții)."),
])
say("Nu toți stakeholderii cer aceeași atenție. Cu unii vorbești zilnic, pe alții doar îi informezi.")
para("Ai mutat pe cineva? De ce utilizatorii au putere mică, deși produsul e făcut pentru ei? Răspuns: individual pot puțin, dar împreună, prin recenzii și abandon, pot mult. Discuția este deschisă.", "Întrebare către sală:")

h2("Slide 18 · Managerul de proiect: tehnic sau om?", "1,5 min")
para("Persoana responsabilă cu gestiunea proiectului și a așteptărilor stakeholderilor.", "Definiție:")
para("Competențele din curs, grupate de noi pe trei familii:", "")
bullets([
    ("Oameni:", "comunicare și negociere, abilități de leader, bun simț și stil."),
    ("Mindset:", "orientare către scop, predispoziție către risc, gândire creativă, corectitudine profesională."),
    ("Tehnic:", "cunoștințe tehnice solide."),
])
say("Observați: din nouă competențe doar una e strict tehnică. Nu înseamnă că tehnicul nu contează, ci că un manager se măsoară mai ales prin oameni și decizii.")
para("POLL 4: Ce manager de proiect alegi? Tehnician genial cu comunicare slabă, sau comunicator excelent cu tehnic mediu? Argumente pentru comunicator: coordonează, negociază, ține stakeholderii aliniați. Argumente pentru tehnician: judecă estimările și riscurile, are credibilitate în echipă. Concluzia bună: depinde de proiect, dar în echipă ai nevoie de ambele.", "")

h2("Slide 19 · Mic exercițiu de gândire (împreună cu A)", "1,5 min")
para("Ce a lipsit cel mai mult în ultimul proiect de echipă care a mers prost? Variante: un plan clar, un stakeholder implicat, comunicarea. Pornești POLL 5, arăți rezultatul și iei 2 comentarii. Mesajul: toate trei sunt probleme de management, nu de cod.")

h2("Slide 20 · Cadrul de gestiune a proiectului", "0,5 min")
para("Cele 8 preocupări din curs, care sunt chiar restul materiei: fezabilitate, scopuri, timp, costuri, schimbări și configurații, calitate, riscuri, resurse umane.")
say("Gantt, CPM și PERT răspund la timp. Stakeholderii leagă scopul de oameni. Brooks arată limita resurselor umane.")

h2("Slide 21 · Ce luăm cu noi", "0,5 min")
bullets([
    "Procesul bate tehnica: cod bun fără proces controlat nu garantează un produs livrat.",
    "Planul e instrument, nu dogmă: ce contează pentru termen se vede în plan.",
    "Oamenii sunt proiectul: stakeholderii și echipa decid reușita mai des decât tehnologia.",
])
para("Apoi: „Întrebări?”. Dacă nu vine nimic, pune tu una: „Cine a avut un stakeholder pe care l-a uitat și l-a costat?”.")

h1("5. Cum rulezi poll-ul (tu și A)")
bullets([
    ("Pe PC:", "deschide într-un tab separat https://mpsprezentare.netlify.app/master și intră cu PIN-ul. Nu îl scrie niciodată pe slide."),
    ("La o întrebare marcată POLL:", "apeși Start lângă ea. Publicul vede întrebarea pe telefon. Aștepți ~30 de secunde."),
    ("Pentru rezultate:", "apeși „Stop + arată rezultate”. Graficul apare pe telefoane și pe panoul tău. Treci pe tab-ul panoului ca să îl arăți sălii, apoi te întorci în PowerPoint."),
    ("Înainte de prezentare:", "fă un test cu telefonul tău și șterge voturile de test cu „Șterge voturile la această întrebare”."),
    ("Dacă internetul cade:", "sari peste poll și întreabă cu mâna sus. Conținutul slide-urilor merge fără net."),
])

h1("6. Întrebări posibile din sală sau de la profesor")
qa("Ce e diferența dintre proiect și program?", "Proiectul produce un rezultat unic. Programul grupează mai multe proiecte înrudite, gestionate coordonat pentru un beneficiu pe care nu l-ai obține gestionându-le separat.")
qa("De ce spunem că proiectul e temporar, dacă software-ul rulează ani?", "Temporar înseamnă că munca de realizare are început și sfârșit. Rezultatul poate dura. După livrare, menținerea este activitate operațională.")
qa("Cine e stakeholder într-un proiect software?", "Oricine e implicat sau afectat: echipa, managerul, clientul, utilizatorii, sponsorul, organizația participantă și externii care pot influența (autorități, comunitate, concurenți). Inclusiv managerul și echipa.")
qa("Poate un stakeholder să fie negativ?", "Da. Poate influența proiectul pozitiv sau negativ. De aceea îi identificăm din timp și îi implicăm, în loc să îi descoperim la final.")
qa("De ce nu e suficient să fii bun tehnic ca să conduci un proiect?", "Pentru că problemele proiectelor sunt de planificare, comunicare, riscuri și așteptări, nu doar de cod. Cursul enumeră nouă competențe, dintre care una singură e strict tehnică.")
qa("Ce înseamnă elaborare progresivă?", "Că proiectul se definește pe măsură ce înaintezi, pe pași și incremente, ținând permanent în vedere scopul final.")
qa("Matricea putere–interes e din curs?", "Nu. E un instrument cunoscut de analiză a stakeholderilor, pe care l-am folosit ca aplicație a ideii din curs.")
qa("Care e diferența dintre drumul critic și Gantt?", "(Răspunde A.) Gantt arată planul în timp. Drumul critic arată care activități nu au voie să întârzie, ca să nu întârzie proiectul.")

h1("7. Glosar rapid")
rows = [
    ("PMBOK", "ghid de bune practici în managementul proiectelor; sursa definițiilor de proiect și stakeholder."),
    ("Stakeholder", "persoană sau organizație implicată în proiect sau afectată de el."),
    ("Sponsor", "cine finanțează proiectul."),
    ("Elaborare progresivă", "definire pe pași și incremente."),
    ("Program", "proiecte înrudite gestionate coordonat."),
    ("Portofoliu", "proiecte/programe fără legătură, grupate strategic."),
    ("ERP", "sistem integrat pentru procesele unei firme."),
    ("One-off", "sistem făcut special pentru un singur client."),
    ("Off-the-shelf", "produs standard pentru mulți utilizatori."),
    ("Integrare orizontală / verticală", "sisteme similare / etape diferite ale unei proceduri."),
    ("Slack", "rezerva de timp a unei activități care nu e pe drumul critic."),
    ("Operațional", "activitate continuă care susține business-ul."),
]
t = doc.add_table(rows=0, cols=2)
t.style = "Table Grid"
for k, v in rows:
    cells = t.add_row().cells
    cells[0].width = Cm(5)
    cells[1].width = Cm(12)
    shade(cells[0], "EEF2F7")
    rr = cells[0].paragraphs[0].add_run(k)
    rr.bold = True
    cells[1].paragraphs[0].add_run(v)

h1("8. Ultimul sfat, în seara dinaintea prezentării")
bullets([
    "Citește secțiunea 4 o dată, apoi spune cu voce tare doar prima propoziție de la fiecare slide, fără să te uiți.",
    "Învață pe de rost definițiile de la slide-urile 10 și 15. Restul îl poți spune cu vorbele tale.",
    "Pregătește un exemplu propriu pentru stakeholderi (un proiect de la facultate) ca să nu te blochezi la exercițiu.",
    "Deschide notele de vorbitor în PowerPoint (Presenter view): fiecare slide are indicații și răspunsuri.",
])

doc.save("/home/user/PREZENTARE_MPS/Ghid_Prezentator_B.docx")
