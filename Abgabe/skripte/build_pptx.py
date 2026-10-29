# -*- coding: utf-8 -*-
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
import data as D
from pptx.oxml.ns import qn


def de(x, nd=2):
    return f"{x:,.{nd}f}".replace(',', 'X').replace('.', ',').replace('X', '.')

OUT = sys.argv[1]
# HTWG-Design (Farben aus den Vorlesungsfolien übernommen)
BG = RGBColor(0xD7, 0xE2, 0xEA)       # Hintergrund hellblau-grau
NAVY = RGBColor(0x33, 0x3F, 0x4F)     # Schiefer (Tabellenkopf, dunkle Flächen)
TEAL = RGBColor(0x00, 0x9A, 0x91)     # HTWG-Petrol (Akzent)
ORANGE = TEAL
GREEN = TEAL
RED = RGBColor(0x10, 0x10, 0x10)
GRAY = RGBColor(0x55, 0x5F, 0x6B)
ROW = RGBColor(0xED, 0xF2, 0xF5)
OK_F = RGBColor(0xCC, 0xEB, 0xE8)
WARN_F = RGBColor(0xF3, 0xE9, 0xC6)
BAD_F = RGBColor(0xF1, 0xD6, 0xD3)
FONT = "Arial"
GRUPPE = "Gruppe: [Name 1, Matr.-Nr.] · [Name 2, Matr.-Nr.] · [Name 3, Matr.-Nr.]"
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = WHITE   # Karten: weiß auf hellblauem Grund
BLACK = RGBColor(0x10, 0x10, 0x10)

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
NW = [sum(c[1] * c[2][j] for c in D.NWA) / 100 for j in range(len(D.SYSTEME))]
T = [D.tco(i) for i in range(len(D.TCO_SYS))]
TN = [T[k] for k in D.NWA_TO_TCO]


def mio(x):
    return f"{x/1e6:.1f}".replace(".", ",")


def box(sl, x, y, w, h, fill=None, line=None):
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    s.shadow.inherit = False
    st = s._element.find(qn('p:style'))
    if st is not None:
        s._element.remove(st)
    return s


def text(sl, x, y, w, h, paras, size=12, color=BLACK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None, line=None, margin=0.05, bar=True):
    """paras: str oder Liste aus str / (str, dict)"""
    s = box(sl, x, y, w, h, fill, line)
    if fill is WHITE and bar:
        box(sl, x, y, 0.06, h, TEAL)
        margin = max(margin, 0.14)
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Inches(margin))
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        opts = {}
        if isinstance(p, tuple):
            p, opts = p
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = opts.get("align", align)
        para.space_after = Pt(opts.get("after", 2))
        r = para.add_run()
        r.text = p
        f = r.font
        f.size = Pt(opts.get("size", size))
        f.bold = opts.get("bold", bold)
        f.color.rgb = opts.get("color", color)
        f.name = FONT
    return s


def slide(title, kicker=None, num=None):
    sl = prs.slides.add_slide(BLANK)
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = BG
    # HTWG-Wortmarke oben rechts
    text(sl, 12.25, 0.22, 0.7, 0.75, [("H   T", {"align": PP_ALIGN.CENTER, "after": 0}), ("W", {"align": PP_ALIGN.CENTER, "after": 0}), ("G", {"align": PP_ALIGN.CENTER, "after": 0})],
         size=11, bold=True, margin=0)
    if kicker:
        text(sl, 0.5, 0.22, 11.5, 0.3, kicker, size=10.5, color=TEAL, bold=True)
    text(sl, 0.5, 0.48, 11.6, 0.9, title, size=23, color=BLACK, bold=True)
    text(sl, 0.5, 7.1, 9, 0.25, "HTWG Konstanz · ERP-Systeme · Fallstudie KVA", size=8, color=GRAY)
    for dx in (10.9, 11.6):
        d = sl.shapes.add_shape(MSO_SHAPE.OVAL, Inches(dx), Inches(7.19), Inches(0.07), Inches(0.07))
        d.fill.solid(); d.fill.fore_color.rgb = BLACK; d.line.fill.background()
        st = d._element.find(qn('p:style'))
        if st is not None:
            d._element.remove(st)
    if num:
        c = sl.shapes.add_shape(MSO_SHAPE.OVAL, Inches(12.45), Inches(6.62), Inches(1.76), Inches(1.76))
        c.fill.solid(); c.fill.fore_color.rgb = TEAL; c.line.fill.background()
        st = c._element.find(qn('p:style'))
        if st is not None:
            c._element.remove(st)
        text(sl, 12.62, 6.92, 0.7, 0.35, str(num), size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    return sl


def source(sl, t, y=6.75):
    text(sl, 0.5, y, 11.7, 0.32, "Quellen: " + t, size=7.5, color=GRAY)


def table(sl, x, y, w, h, rows, colw, size=10, header=True, fills=None):
    shp = sl.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h))
    tb = shp.table
    for j, cw in enumerate(colw):
        tb.columns[j].width = Inches(cw)
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            c = tb.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.05)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.text = str(v)
            for p in c.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)
                    run.font.name = FONT
                    run.font.bold = header and i == 0
                    run.font.color.rgb = WHITE if header and i == 0 else BLACK
            c.fill.solid()
            if header and i == 0:
                c.fill.fore_color.rgb = NAVY
            elif fills and (i, j) in fills:
                c.fill.fore_color.rgb = fills[(i, j)]
            else:
                c.fill.fore_color.rgb = WHITE if i % 2 else ROW
    return tb


# ============================================================ FOLIE 1
sl = slide("Warum jetzt? Fünf Fristen bis Ende 2027 – KVA-WARE hält keine davon",
           "1 · AUSGANGSLAGE UND HANDLUNGSDRUCK", 1)
text(sl, 0.5, 1.0, 7.7, 0.36, GRUPPE, size=10.5, bold=True, color=WHITE, fill=NAVY, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
# Zeitstrahl
box(sl, 0.6, 2.45, 7.6, 0.05, NAVY)
events = [
    ("31.12.2026", "NAV-Partner Pilsen kündigt", "Server 2016, kein Failover – Werk liefert 40 % der Gehäuse"),
    ("01.01.2027", "Pflicht E-Rechnungs-Versand", "gilt für KVA schon 2027 (Umsatz > 800 T€) – nicht erst 2028"),
    ("2027", "BDE-Terminals am Ende", "Ersatzteile nur noch gebraucht (eBay)"),
    ("03/2027", "Salesforce-Vertrag endet", "nur noch 1 Jahr Verlängerung ohne Mehrkosten"),
    ("31.12.2027", "Ruhestand K. Brugger", "einziger KVA-WARE-Entwickler – „Das Wichtigste ist im Kopf“"),
]
for i, (d, t1, t2) in enumerate(events):
    x = 0.6 + i * 1.55
    s = sl.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.55), Inches(2.36), Inches(0.22), Inches(0.22))
    s.fill.solid(); s.fill.fore_color.rgb = RED if i in (0, 1, 4) else ORANGE; s.line.fill.background()
    text(sl, x, 1.6, 1.5, 0.7, [(d, {"bold": True, "size": 14, "color": RED if i in (0, 1, 4) else ORANGE}), (t1, {"bold": True, "size": 10})], size=10)
    text(sl, x, 2.7, 1.5, 1.0, t2, size=9.5, color=GRAY)
text(sl, 0.6, 3.6, 3.7, 1.05,
     [("Auftrag Beirat (Q3/2025)", {"bold": True, "size": 10.5, "color": NAVY}),
      ("Abhängigkeit von KVA-WARE bis Ende 2027 lösen; Vorlage mit ≥ 3 Alternativen bis Q4/2026. Meilenstein „Vertrag 30.09.2026“ ist schon verstrichen.", {"size": 9.5})], fill=LIGHT, margin=0.08)
text(sl, 4.45, 3.6, 3.75, 1.05,
     [("Gescheitert: „KVA neu“ 2014–2016", {"bold": True, "size": 10.5, "color": RED}),
      ("SAP-Projekt nach 17 Monaten abgebrochen, 940 T€ ohne Ergebnis. Ursachen: kein Konfigurator-Konzept, Key User nicht frei, Stammdaten, revidierte Entscheidungen.", {"size": 9.5})],
     fill=WHITE, margin=0.08)
# KPIs
text(sl, 8.7, 1.5, 4.2, 0.35, "Strategie 2030: heute nicht erreichbar", size=13, bold=True, color=NAVY)
kpis = [("Liefertermintreue", "58 %*", "> 92 %"), ("Durchlaufzeit Standard", "18 AT", "10 AT"), ("Lagerbestand", "18,9 Mio. €", "12 Mio. €"),
        ("Monatsabschluss Gruppe", "14 AT", "5 AT"), ("Wartungsverträge", "41", "300"), ("EBIT-Marge", "4,4 %", "9 %")]
table(sl, 8.7, 1.9, 4.2, 2.4, [("Kennzahl", "2025", "Ziel")] + kpis, [2.0, 1.1, 1.1], size=10.5)
text(sl, 8.7, 4.35, 4.2, 0.4, "* gegen ersten bestätigten Termin; KVA-WARE misst gegen den letzten (71 %)", size=8.5, color=GRAY)
# Ursachen
text(sl, 0.6, 4.75, 12.3, 0.35, "Kern des Problems: Insellösungen statt durchgängigem Beleg", size=13, bold=True, color=NAVY)
causes = [
    ("Kein Konfigurator im ERP", "Excel-Tool „KonfiX“ (2009, ohne Betreuer), 38.000 Einzel-Stücklisten, 60 % nie wieder genutzt"),
    ("Termine ohne Grundlage", "keine Kapazitätsplanung, Härterei per Magnettafel, 35 % der FA starten mit Fehlteilen"),
    ("Keine Rückverfolgbarkeit", "keine Seriennummern/Chargen – Rückruf 2024 kostete 180 T€; Installed Base (60.000 Antriebe) unbekannt"),
    ("Drei Systemwelten", "KVA-WARE (nur EUR, nur DE), NAV Pilsen, Bexio CH – Konsolidierung in Excel"),
]
for i, (h, t) in enumerate(causes):
    text(sl, 0.6 + i * 3.1, 5.15, 2.95, 1.3, [(h, {"bold": True, "size": 11.5, "color": NAVY}), (t, {"size": 10})], fill=LIGHT, margin=0.1)
source(sl, "Fallpaket KVA: Projektauftrag v1.2; IT-Architektur IST v0.9; Controlling 2025; Interviews 04, 06, 09–11; Beiratsprotokoll Q3/2025; Mail Nováková 04.02.2026; Helpdesk T25-0412; E-Rechnung: IHK Frankfurt/BMF.")

# ============================================================ FOLIE 2
sl = slide("Aus 111 Rohanforderungen wird ein Kern aus 8 K.o.-Kriterien – fünf Zielkonflikte sind aufzulösen",
           "2 · ZENTRALE ANFORDERUNGEN, K.-O.-KRITERIEN UND ZIELKONFLIKTE", 2)
text(sl, 0.5, 1.35, 5.4, 0.35, "K.-o.-Kriterien (Ausschluss bei Nichterfüllung)", size=13, bold=True, color=NAVY)
ko_rows = [(k[0], k[1]) for k in D.KO]
table(sl, 0.5, 1.75, 5.4, 3.3, [("", "Kriterium")] + ko_rows, [0.5, 4.9], size=10.5)
text(sl, 0.5, 5.15, 5.4, 1.5,
     [("Konsolidierung (Artefakt B)", {"bold": True, "color": NAVY, "size": 11}),
      ("111 Rohanforderungen + 8 neue → 14 K.o. · 61 Muss · 27 Kann · 5 Phase 2 · 6 gestrichen", {"size": 10}),
      ("Lösungsvorgaben umformuliert, z. B. „Ansicht wie F12“ → „Arbeitsplatzübersicht mit Ampel, < 2 s“; „alle 340 Listen“ → Top-40 lt. Nutzungsstatistik", {"size": 10}),
      ("42 Widersprüche mit Fundstelle dokumentiert (Anhang A1)", {"size": 10})], fill=LIGHT, margin=0.1)
text(sl, 6.2, 1.35, 6.7, 0.35, "Zielkonflikte und unser Lösungsvorschlag", size=13, bold=True, color=NAVY)
zk = [
    ("Konflikt", "Positionen", "Vorschlag"),
    ("Seriennummern", "Service/QM: alle Antriebe – Lager/GF Technik: nur Sonder", "Alle Antriebe, automatisch beim Typenschilddruck, Scan statt Abtippen"),
    ("Betriebsmodell", "CFO/CEO: SaaS – IT/GF Technik: Private Cloud/On-Prem", "Kein K.o.: „SaaS mit Plan B“ (EU-RZ, Exit, Offline-Puffer, 2. Leitung)"),
    ("Pilsen", "Projektauftrag: 2029 – Werk/CFO: sofort", "Pilsen zuerst als Pilot (sequenziell), NAV-Überbrückung 9 Monate"),
    ("CRM", "CEO: ein System – Vertrieb: „Salesforce bleibt“", "Salesforce bis 03/2028 verlängern + integrieren, CRM-Entscheid 2028"),
    ("Rückmeldung", "Controlling: je Personalnummer – Betriebsrat: lehnt ab", "Arbeitsplatzbezogen + Lohngruppen-Mischsatz, Rahmen-BV vor Vertrag"),
]
table(sl, 6.2, 1.75, 6.7, 4.0, zk, [1.25, 2.6, 2.85], size=10)
text(sl, 6.2, 5.85, 6.7, 0.8,
     [("Ausland beachtet: CZ = EU (Intrastat, Kontrollmeldung, CZ-Sprache) · CH = Drittland (Zoll bei jeder Lieferung, MWST, QR-Rechnung, CHF) → K3/K4, Anhang A10", {"size": 9.5}),
      ("20 Störinformationen bewusst aussortiert (veraltetes Lastenheft 2015, Anekdoten, iPad-3D-Wunsch, falsche Zahlen) → Anhang A9", {"size": 9.5, "color": GRAY})])
source(sl, "Anforderungskatalog Rohfassung 12.02.2026; Interviews 01–12; Flipcharts 21.01.2026; Kick-off 02.12.2025; LK-Ergebnis 18.02.2026; BR-Schreiben 09.02.2026. Details: XLSX Blätter 1–3.")

# ============================================================ FOLIE 3
sl = slide("Markt: 16 Optionen geprüft – vier erfüllen alle K.o.-Kriterien und passen zu Größe und Zeitplan",
           "3 · MARKTÜBERBLICK: LONGLIST → SHORTLIST", 3)
stages = [
    (16, "Longlist", "Mittelstands-ERP für diskrete Fertiger in DACH/CZ + Option 0 (Modernisierung KVA-WARE)", 6.0),
    (14, "Filter K1/K2: Konfigurator + Projektfertigung", "raus: Option 0 (KVA-WARE modernisieren), Odoo", 5.2),
    (8, "Filter K4/K6: Lokalisierung + Partner DE UND CZ", "raus: Infor LN, Epicor, PSIpenta, ams.erp, HELIOS (nur CZ); Nachrücker: APplus (kein APplus-Partner in CZ belegt)", 4.4),
    (4, "Filter Größe, Kosten, Betriebsmodell, Akzeptanz", "raus: SAP S/4 Public (nur SaaS, kein E-Invoicing CZ), SAP S/4 Private + D365 SCM (über Rahmen), Infor CSI (US-Cloud ohne Plan B)", 3.6),
]
for i, (n, h, t, w) in enumerate(stages):
    y = 1.45 + i * 1.02
    x = 0.5 + (6.0 - w) / 2
    text(sl, x, y, w, 0.88, [(f"{n}", {"bold": True, "size": 20, "color": WHITE, "align": PP_ALIGN.CENTER}),
                             (h, {"size": 9.5, "color": WHITE, "align": PP_ALIGN.CENTER})],
         fill=NAVY if i < 3 else TEAL, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, 6.7, y + 0.03, 6.2, 0.85, t, size=10, anchor=MSO_ANCHOR.MIDDLE)
cards = [
    ("Business Central + Branchenlösung", "Abo je Nutzer · SaaS, Hosting als Plan B · CZ-Lokalisierung von Microsoft · Mehrmandanten", "S02, S03"),
    ("proALPHA", "Kauf oder Abo · Spezialist Variantenfertigung · CZ: SPC solutions (Zlín) · Mehrmandanten", "S05, S06"),
    ("IFS Cloud", "Abo · ETO/Service-stark · EU-Anbieter, SaaS oder kundenbetrieben · Multi-Site", "S09–S11"),
    ("abas ERP", "Kauf oder Abo · Fertigungs-Mittelstand · CZ: amotIQ · Mehrmandanten", "S07, S08"),
]
for i, (h, t, q) in enumerate(cards):
    text(sl, 0.5 + i * 3.1, 5.6, 2.95, 1.1, [(f"{i+1}  {h}", {"bold": True, "size": 10.5, "color": NAVY}), (t, {"size": 8.5}), (q, {"size": 7.5, "color": GRAY})], fill=WHITE, margin=0.08)
source(sl, "S02–S16, Abruf 29.10.2026; Systemarchitektur (Multi-Site/Mehrmandanten) nach Kurz (2026) und Martin/Kurz (2024) = S20, S21. Marktanteilsstudien bewusst nicht genutzt („Datenquellen meist zweifelhaft“, S21). Longlist komplett: Anhang A7.")

# ============================================================ FOLIE 4
sl = slide("IFS bietet den höchsten Nutzen – Business Central das beste Verhältnis von Nutzen zu Kosten",
           "4 · BEWERTUNGSERGEBNIS: NUTZWERT UND TCO ÜBER 5 JAHRE", 4)
cd = CategoryChartData()
cd.categories = D.SYSTEME
cd.add_series("Nutzwert (max. 5,0)", [round(v, 2) for v in NW])
gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.45), Inches(4.0), Inches(3.3), cd)
ch = gf.chart
ch.has_legend = False
ch.font.name = FONT
ch.has_title = True
ch.chart_title.text_frame.text = "Nutzwert (gewichtet, max. 5,0)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
ch.value_axis.maximum_scale, ch.value_axis.minimum_scale = 5.0, 0
ch.value_axis.has_major_gridlines = False
ch.value_axis.tick_labels.font.size = Pt(9)
ch.category_axis.tick_labels.font.size = Pt(9)
pl = ch.plots[0]
pl.has_data_labels = True
pl.data_labels.number_format, pl.data_labels.number_format_is_linked = '0.00', False
pl.data_labels.font.size, pl.data_labels.font.bold = Pt(11), True
pl.series[0].format.fill.solid(); pl.series[0].format.fill.fore_color.rgb = NAVY

cd2 = CategoryChartData()
cd2.categories = ["BC + Branchen-lösung", "BC Basis (günstig)", "proALPHA", "IFS Cloud", "abas"]
cd2.add_series("Subskription", [round(t["sub5"] / 1e6, 2) for t in T])
cd2.add_series("Einführung inkl. Puffer", [round(t["einm5"] / 1e6, 2) for t in T])
cd2.add_series("Betrieb", [round(t["betr5"] / 1e6, 2) for t in T])
gf2 = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(4.7), Inches(1.45), Inches(4.4), Inches(3.3), cd2)
c2 = gf2.chart
c2.has_title = True
c2.font.name = FONT
c2.chart_title.text_frame.text = "TCO 2027–2031 in Mio. €"
c2.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
c2.has_legend = True
c2.legend.position = XL_LEGEND_POSITION.BOTTOM
c2.legend.include_in_layout = False
c2.legend.font.size = Pt(8.5)
c2.value_axis.has_major_gridlines = False
c2.value_axis.maximum_scale = 4.0
c2.plots[0].gap_width = 60
c2.value_axis.tick_labels.font.size = Pt(9)
c2.category_axis.tick_labels.font.size = Pt(8.5)
for s, col in zip(c2.plots[0].series, [NAVY, ORANGE, RGBColor(0x9A, 0xA8, 0xB5)]):
    s.format.fill.solid(); s.format.fill.fore_color.rgb = col
r2 = lambda x: f"{x:.2f}".replace(".", ",")
rows = [("", "BC", "proALPHA", "IFS", "abas"),
        ("Nutzwert", *[r2(v) for v in NW]),
        ("TCO 5 J. (Mio. €)", *[mio(t["total"]) for t in TN]),
        ("Spanne (Mio. €)", *[f"{mio(t['low'])}–{mio(t['high'])}" for t in TN]),
        ("Nutzwert je Mio. €", *[r2(NW[j] / TN[j]["total"] * 1e6) for j in range(4)]),
        ("Pilsen bis Q3/2027", "realistisch", "knapp", "knapp", "knapp"),
        ("Hauptrisiko", "Konfigurator", "CZ-Lokal.", "Kosten", "CZ-Lokal.")]
fills = {(4, 1): OK_F, (1, 3): OK_F, (5, 1): OK_F}
table(sl, 9.2, 1.5, 3.75, 3.25, rows, [1.05, 0.7, 0.7, 0.6, 0.7], size=8)
text(sl, 0.5, 4.95, 6.1, 1.75,
     [("So lesen wir das Ergebnis", {"bold": True, "color": NAVY, "size": 12}),
      ("• IFS liegt vorn (ETO, Service, Multi-Site-Architektur), kostet aber rund 1,2 Mio. € mehr als BC; proALPHA ist stabil Platz 2.", {"size": 10}),
      ("• BC: bester Nutzen je Euro, schnellster Weg für Pilsen, Hosting als Plan B; bei „Fokus Zeit & Pilsen“ nahe an IFS (A3).", {"size": 10}),
      ("• abas ist günstiger als proALPHA, fällt aber bei Service und CZ-Lokalisierung zurück.", {"size": 10}),
      ("• Vollkosten je Nutzer und Monat: BC ≈ 186 €, abas ≈ 212 €, proALPHA ≈ 230 €, IFS ≈ 275 €.", {"size": 10})])
text(sl, 6.8, 4.95, 6.1, 1.75,
     [("Budget: keine Option passt in 1,2 Mio. € oder < 1 Mio. €", {"bold": True, "color": RED, "size": 12}),
      ("• Interne Kosten im Projektauftrag 120 T€ statt realistisch 350–450 T€. Selbst in seiner Logik (Einführung + 3 J. Lizenzen) kostet BC ≈ 1,9 Mio. €.", {"size": 10}),
      (f"• Auch die günstige Variante (BC Basis, {mio(T[1]['total'])} Mio. €) liegt über der CFO-Grenze von 1,5 Mio. €.", {"size": 10}),
      ("• Einordnung: 0,5–0,9 % vom Umsatz p. a.; Nutzen 0,5–0,9 Mio. € p. a. + 3–4 Mio. € Liquidität aus Bestandsabbau.", {"size": 10})])
source(sl, "Eigene Nutzwertanalyse (11 Kriterien, Gewichte vorab festgelegt) und TCO-Modell, XLSX Blätter 6–9. BC-Preise lt. Microsoft-Preisliste DE (S02); proALPHA/IFS/abas-Preise nicht öffentlich = gekennzeichnete Annahmen, Spanne −15 %/+35 %.")

# ============================================================ FOLIE 5
sl = slide("Empfehlung: Business Central + Branchenlösung als Favorit – Vertrag erst nach bestandenem Konfigurator-Test",
           "5 · EMPFEHLUNG, ROADMAP UND ENTSCHEIDUNGSBEDARF DER GESCHÄFTSFÜHRUNG", 5)
# Roadmap
quarters = ["Q4/26", "Q1/27", "Q2/27", "Q3/27", "Q4/27", "Q1/28", "Q2/28", "H2/28", "2029"]
x0, wq, y0 = 3.55, 0.52, 1.4
text(sl, 0.5, y0, 3.0, 0.3, "Roadmap", size=12, bold=True, color=NAVY)
for i, q in enumerate(quarters):
    text(sl, x0 + i * wq, y0, wq, 0.28, q, size=8, color=GRAY, align=PP_ALIGN.CENTER)
bars = [
    ("Sofortmaßnahmen", "E-Rechnung-Konverter, NAV-Notfallsupport, Brugger-Doku, Stammdaten", 0, 2, ORANGE, "ab sofort"),
    ("Ausschreibung + PoC", "RFP an 4 Anbieter, gleiches Testskript (A11), PoC BC vs. proALPHA", 0.15, 2, NAVY, "Vertrag 04/27"),
    ("Pilsen (Pilot)", "Template für die Gruppe, NAV-Brücke bis 09/2027", 2, 4, GREEN, "Go-live 01.10.27"),
    ("Singen + CH", "Kern: Konfigurator, Fertigung, Finanzen, QM, Service", 2.5, 6, NAVY, "Go-live 01.04.28"),
    ("Phase 2", "APS, Field-Service-App, Portal/Shop, Instandhaltung, CRM", 6, 9, GRAY, "ab H2/28"),
]
for k, (lab, det, s0, e, col, tip) in enumerate(bars):
    y = y0 + 0.32 + k * 0.48
    text(sl, 0.5, y, 3.0, 0.46, [(lab, {"bold": True, "size": 9.5}), (det, {"size": 7.5, "color": GRAY})], margin=0.02)
    text(sl, x0 + s0 * wq, y + 0.06, (e - s0) * wq, 0.32, tip, size=8, bold=True, color=WHITE, fill=col, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.02)
# Was heißt das für Sie (Management abholen)
text(sl, 0.5, 4.05, 7.6, 0.3, "Was die Empfehlung für Sie bedeutet", size=12, bold=True, color=NAVY)
srow = [("Wer", "Ihr Anliegen", "Unsere Antwort")] + [r for r in D.STAKE if r[0] != "Beirat"]
table(sl, 0.5, 4.32, 7.6, 1.9, srow, [1.75, 2.75, 3.1], size=7)
text(sl, 0.5, 6.42, 7.6, 0.3, "Bedingung (Lehre aus 2016): Vertrag erst nach bestandenem Konfigurator-PoC – sonst proALPHA (+0,6 Mio. €).",
     size=9, bold=True, color=RED)
# Entscheidungen GF
text(sl, 8.35, 1.4, 4.55, 5.3,
     [("Das muss die Geschäftsführung entscheiden", {"bold": True, "color": WHITE, "size": 12.5}),
      ("1  Budget: 5-J.-TCO bis 3,0 Mio. € freigeben (BC 2,6 Mio. € + Reserve) statt 1,2 Mio. €", {"color": WHITE, "size": 10.5, "after": 5}),
      ("2  Kosten/Tempo oder Funktionstiefe: BC (Favorit) oder proALPHA/IFS – Entscheidung nach dem PoC im März 2027", {"color": WHITE, "size": 10.5, "after": 5}),
      ("3  Termin: Go-live Singen 01.04.2028 statt 01.01.2028; Bereitschaft Brugger + RPG-Dienstleister als Brücke", {"color": WHITE, "size": 10.5, "after": 5}),
      ("4  Pilsen zuerst: NAV-Überbrückung mit DataSoft bis 09/2027 verhandeln", {"color": WHITE, "size": 10.5, "after": 5}),
      ("5  Seriennummer für alle Antriebe – endgültig, Änderung nur per Lenkungskreis", {"color": WHITE, "size": 10.5, "after": 5}),
      ("6  Salesforce 1 Jahr verlängern; CRM-Entscheidung 2028", {"color": WHITE, "size": 10.5, "after": 5}),
      ("7  Key User zu 20–30 % freistellen, externer Projektleiter, Exportkontrollbeauftragten benennen", {"color": WHITE, "size": 10.5, "after": 5}),
      ("8  E-Rechnungs-Konverter für KVA-WARE sofort beauftragen (Pflicht ab 01.01.2027)", {"color": WHITE, "size": 10.5})],
     fill=NAVY, margin=0.15)
source(sl, "Eigene Empfehlung auf Basis XLSX Blätter 1–9; Termine aus Mail Nováková 04.02.2026, Int. 04 (Brenner), LK 18.02.2026; Nächste Schritte bis Vertrag: Anhang A6.")

# ============================================================ ANHANG
sl = slide("Anhang A1 – Wichtigste Widersprüche (alle 42 in XLSX Blatt 1)", "ANHANG", "A1")
top = [w for w in D.WIDERSPRUECHE if w[8] == "hoch"][:13]
rows = [("ID", "Thema", "Aussage A", "Aussage B", "Lösungsvorschlag")] + [(w[0], w[1], w[2], w[3], w[6]) for w in top]
table(sl, 0.4, 1.3, 12.55, 5.6, rows, [0.5, 1.15, 3.3, 3.6, 4.0], size=7)

sl = slide("Anhang A2 – K.o.-Prüfung: Shortlist und Option 0", "ANHANG", "A2")
ko = [("K.o.-Kriterium", "BC + Branchenlösung", "proALPHA", "IFS Cloud", "abas ERP", "Option 0")]
vals = {
    "K1": ("✓ über ISV – PoC", "✓ lt. Hersteller", "✓", "✓ PoC", "✗"), "K2": ("✓ (Projekte, einfacher)", "✓", "✓", "✓", "✗ (Excel)"),
    "K3": ("✓ (Mehrmandanten + IC)", "✓ (Mehrmandanten)", "✓ (Multi-Site)", "✓ (Mehrmandanten)", "✗"),
    "K4": ("✓ (Microsoft CZ)", "? im RFI", "✓", "? im RFI", "✗"),
    "K5": ("✓", "✓", "✓", "✓", "✗"), "K6": ("✓", "✓ SPC solutions", "✓", "✓ amotIQ", "✗"),
    "K7": ("✓ EU-RZ / Hosting DE", "✓ DE", "✓ EU", "✓ DE", "✓ On-Prem"), "K8": ("✓", "✓", "✓", "✓", "teilweise"),
}
fills = {}
for i, k in enumerate(D.KO, 1):
    v = vals[k[0]]
    ko.append((f"{k[0]} {k[1]}",) + v)
    for j, x in enumerate(v, 1):
        fills[(i, j)] = OK_F if x.startswith("✓") else WARN_F if x.startswith("?") or x.startswith("teil") else BAD_F
table(sl, 0.5, 1.4, 12.3, 4.2, ko, [3.7, 2.1, 1.7, 1.6, 1.7, 1.5], size=9.5, fills=fills)
text(sl, 0.5, 5.8, 12.3, 0.9, "Option 0 (RPG→Java-Migration, 300–400 T€ lt. GF Technik) wurde bewusst geprüft: Sie verlagert die Abhängigkeit nur auf einen Dienstleister "
     "und erfüllt 5 von 8 K.o.-Kriterien nicht. Nachrücker bei Ausfall eines Kandidaten: APplus (APplus-Partner in CZ im RFI klären).", size=10.5, fill=LIGHT, margin=0.1)

sl = slide("Anhang A3 – Nutzwertanalyse im Detail und Sensitivität", "ANHANG", "A3")
rows = [("Kriterium", "Gew. %", "BC", "proALPHA", "IFS", "abas")] + [(c[0], c[1], *c[2]) for c in D.NWA] + [("Nutzwert", 100, *[de(v) for v in NW])]
table(sl, 0.4, 1.3, 7.7, 5.2, rows, [4.5, 0.6, 0.6, 0.8, 0.6, 0.6], size=8.5)
srows = [("Szenario", "BC", "proALPHA", "IFS", "abas")]
for name, w in D.SENS.items():
    srows.append((name,) + tuple(de(sum(wi*c[2][j] for wi, c in zip(w, D.NWA))/100) for j in range(4)))
table(sl, 8.3, 1.3, 4.65, 2.0, srows, [1.95, 0.6, 0.85, 0.6, 0.65], size=8.5)
text(sl, 8.3, 3.45, 4.6, 3.0,
     [("Methodik", {"bold": True, "color": NAVY}),
      ("• Gewichte vor der Bewertung festgelegt, abgeleitet aus Umsatzanteilen (CTO 60 %) und Fristen.", {"size": 9.5}),
      ("• Skala 1–5: 3 = nur mit Partner-Erweiterung, 5 = Standard mit Marktstärke.", {"size": 9.5}),
      ("• Neues Kriterium Systemarchitektur nach Vorlesung: Multi-Site führt IC-Aufträge wie interne Aufträge, Mehrmandanten über Belegaustausch (S20, S21).", {"size": 9.5}),
      ("• Kosten bewusst NICHT im Nutzwert, sondern getrennt als TCO → Kennzahl Nutzwert je Mio. €.", {"size": 9.5}),
      ("• Punkte beruhen auf Herstellerangaben und Sekundärquellen – in Drehbuch-Demos zu bestätigen.", {"size": 9.5})], size=9.5)

sl = slide("Anhang A4 – TCO-Modell: Annahmen offengelegt", "ANHANG", "A4")
rows = [("Parameter", "Einheit", "BC", "BC Basis", "proALPHA", "IFS", "abas", "Quelle / Annahme")]
for n, u, v, s in D.TCO_PARAM:
    fmt = (lambda x: f"{round(x*100)} %") if u == "%" else (lambda x: de(x, 2)) if u == "Faktor" else (lambda x: de(x, 0) if float(x).is_integer() else de(x, 1))
    rows.append((n, u, *[fmt(x) for x in v], s))
rows.append(("TCO 5 Jahre (Spanne)", "Mio. €", *[f"{mio(t['total'])} ({mio(t['low'])}–{mio(t['high'])})" for t in T], "Summe 2027–2031"))
table(sl, 0.4, 1.2, 12.55, 5.6, rows, [2.9, 1.1, 0.7, 0.7, 0.75, 0.7, 0.7, 5.0], size=6.5)

sl = slide("Anhang A5 – Risiken und Gegenmaßnahmen (Lehren aus 2016)", "ANHANG", "A5")
risks = [("Risiko", "Ursache 2016 / Fundstelle", "Gegenmaßnahme", "Verantwortlich"),
         ("Konfigurator nicht abbildbar", "3 Konzepte, keines umgesetzt (IT-Historie)", "PoC mit 3 realen Typenschlüsseln VOR Vertrag; Regelwerk-Owner Konstruktion (Dr. Weber)", "LK"),
         ("Key User nicht verfügbar", "Auftragsspitze 2015; Int. 05: 25 % Doppelerfassung", "20–30 % Freistellung schriftlich, Vertretung, Q4-Spitze im Plan meiden", "GF"),
         ("Stammdatenqualität", "Dubletten, keine Revisionen; 38.000 Stücklisten", "Bereinigung ab sofort, nur verwendete Stücklisten migrieren", "IT/AV"),
         ("Revidierte Entscheidungen", "GF Technik revidierte u. a. Seriennummern", "Entscheidungslog, Änderung nur per LK-Beschluss; Vogt unterschreibt Projektauftrag", "GF"),
         ("Unerfüllbare Abnahme", "340 Listen als Abnahmekriterium", "Top-40-Liste aus Nutzungsstatistik, Rest Power BI", "PL"),
         ("Wissensverlust Brugger", "Versteckte Preislogik", "Dokumentation bis 06/2027, Bereitschaft auf Stundenbasis, RPG-Dienstleister", "IT"),
         ("Ausfall NAV Pilsen vor Go-live", "Server 2016, Partner-Kündigung", "Notfallsupport bis 09/2027, VM-Backup des NAV-Servers", "Werk Pilsen"),
         ("Betriebsrat blockiert", "BDE 2012: Einigungsstelle", "Rahmen-BV parallel zum RFP, BR in Demos, keine Personalnummer je AG", "GF/BR"),
         ("Internetausfall (SaaS)", "2 Ausfälle 2025 > 4 h", "2. Leitung (18 T€/J.), Offline-Puffer für Rückmeldung/Versand", "IT")]
table(sl, 0.4, 1.3, 12.55, 5.3, risks, [2.4, 3.3, 5.45, 1.4], size=9.5)

sl = slide("Anhang A6 – Nächste Schritte bis zur Vertragsunterschrift (April 2027)", "ANHANG", "A6")
steps = [("Wann", "Schritt", "Ergebnis"),
         ("Dez 2026", "Beirat: Entscheidungsvorlage, Budgetrahmen, Termine (Folie 5)", "Mandat für Ausschreibung"),
         ("Dez 2026", "Sofortmaßnahmen: E-Rechnungs-Konverter, NAV-Notfallsupport, Brugger-Nutzungsstatistik + Doku, Stammdaten", "Fristen 2027 entschärft"),
         ("Jan 2027", "Lastenheft finalisieren (Basis XLSX Blatt 3), RFP an 4 Anbieter + Partner DE/CZ; CZ-Lokalisierung proALPHA/abas schriftlich belegen", "Vergleichbare Angebote"),
         ("Feb 2027", "Drehbuch-Demos (je 2 Tage) mit identischem Testskript T1–T6 (Anhang A11); Meister + 2 BR-Mitglieder bewerten", "Punkte der Nutzwertanalyse bestätigt"),
         ("Mär 2027", "PoC Konfigurator (2–3 Wochen, BC + proALPHA), Referenzbesuche DE und CZ, Festpreis-/Budgetangebote", "Go/No-Go je Kandidat"),
         ("Mär 2027", "Rahmen-Betriebsvereinbarung IT abschließen; Salesforce 1 Jahr verlängern", "Voraussetzung Vertrag"),
         ("Apr 2027", "Vertragsverhandlung (Exit-Klausel, AVV, SLA, Preisbindung 5 J.), LK-Beschluss, Unterschrift", "Projektstart Pilsen-Pilot")]
table(sl, 0.4, 1.3, 12.55, 4.6, steps, [1.2, 8.0, 3.35], size=10.5)
text(sl, 0.4, 6.05, 12.55, 0.65, "Offene Rückfragen an die GF (Rollenspiel): Status Rahmen-BV? · Gilt 01.01.2028 auch bei Verschiebung des Vertrags? · Ist die Obergrenze 1,5 Mio. € TCO beschlossen oder Wunsch? · "
     "NIS2-Registrierung erfolgt? · Ist DataSoft zu befristetem Notfallsupport bereit?", size=10, fill=LIGHT, margin=0.1)

sl = slide("Anhang A7 – Longlist komplett: wer warum ausscheidet (für Rückfragen)", "ANHANG", "A7")
rows = [("Nr", "Produkt", "Betrieb", "Präsenz CZ", "Ergebnis", "Begründung")] + [(l[0], l[2], l[4], l[6], l[8], l[9]) for l in D.LONGLIST]
fills = {}
for i, l in enumerate(D.LONGLIST, 1):
    fills[(i, 4)] = OK_F if l[8].startswith("Shortlist") else WARN_F if l[8].startswith("Nachr") else BAD_F
table(sl, 0.35, 1.2, 12.65, 5.75, rows, [0.35, 2.6, 1.9, 2.4, 1.2, 4.2], size=6.8, fills=fills)

sl = slide("Anhang A8 – Lizenzmodelle: was wird eigentlich bezahlt?", "ANHANG", "A8")
rows = [("System", "Lizenzmodell", "Nutzerarten", "Shopfloor/Lager", "Kauf möglich?", "Konsequenz für KVA", "Prüfstatus")] + list(D.LIZENZ)
table(sl, 0.35, 1.25, 12.65, 4.3, rows, [1.6, 2.3, 2.1, 1.6, 1.2, 2.45, 1.4], size=8)
text(sl, 0.35, 5.75, 12.65, 0.95,
     [("Warum das wichtig ist: 210 Mitarbeitende melden in der Fertigung zurück. Würde jede Person eine Volllizenz brauchen, stiege die Subskription um ein Vielfaches. "
       "Deshalb rechnen wir mit 120 Voll- und 110 Light-Usern plus 30 Geräte-Lizenzen und fragen das Lizenzmodell für alle drei Nutzerarten im RFP verbindlich ab "
       "(inkl. Preisbindung über 5 Jahre).", {"size": 10.5})], fill=LIGHT, margin=0.1)

sl = slide("Anhang A9 – Störinformationen: was wir bewusst nicht verwendet haben", "ANHANG", "A9")
rows = [("Nr", "Information", "Fundstelle", "Warum irreführend", "Umgang")] + [(x[0], x[1], x[2], x[3], x[4]) for x in D.STOER]
table(sl, 0.35, 1.2, 12.65, 5.75, rows, [0.35, 3.9, 1.8, 3.3, 3.3], size=6.8)

sl = slide("Anhang A10 – Länder und Rechtsrahmen: Inland, EU-Ausland, Drittland", "ANHANG", "A10")
rows = [("Gesellschaft", "Status", "Rechtliche Anforderungen an das ERP", "Weitere Anforderungen")] + list(D.LAENDER)
table(sl, 0.35, 1.3, 12.65, 4.2, rows, [2.0, 1.4, 5.5, 3.75], size=9.5)
text(sl, 0.35, 5.7, 12.65, 1.0,
     [("Folge für die Auswahl: Lokalisierung für DE, CZ und CH ist K.o.-Kriterium K4, mehrere Gesellschaften und Währungen K3. Schweiz-Lieferungen brauchen heute Papier-Zolldokumente "
       "(Interview 08) – Zoll/ATLAS-Anbindung ist Muss (L-006). Datenhaltung: EU-Rechenzentrum ist K7; die Schweiz erkennt das EU-Datenschutzniveau als angemessen an.", {"size": 10.5})], fill=LIGHT, margin=0.1)

sl = slide("Anhang A11 – Ein Testskript für alle Anbieter (Drehbuch-Demo und PoC)", "ANHANG", "A11")
for i, (h, t, g) in enumerate(D.TESTS):
    x = 0.5 + (i % 2) * 6.25
    y = 1.35 + (i // 2) * 1.75
    text(sl, x, y, 6.05, 1.55, [(h, {"bold": True, "size": 12, "color": NAVY}), (t, {"size": 10}), (g, {"size": 9, "color": TEAL, "bold": True})], fill=WHITE, margin=0.1)
text(sl, 0.5, 6.62, 11.7, 0.3, "Für jeden Test: Daten vorab liefern, Ergebnis dokumentieren, Lücken/Add-ons/Partnerleistung getrennt bepreisen. Ein nicht bestandener K.o.-Test schlägt den Nutzwert.", size=9.5, bold=True)

sl = slide("Anhang A12 – Quellen", "ANHANG", "A12")
rows = [("ID", "Quelle", "Art / Prüfstatus")] + [(q[0], q[1], q[2]) for q in D.QUELLEN]
table(sl, 0.35, 1.15, 12.65, 5.8, rows, [0.5, 10.0, 2.15], size=6.3)

for i, sl_ in enumerate(prs.slides, 1):
    if i in D.NOTES:
        sl_.notes_slide.notes_text_frame.text = D.NOTES[i]
prs.save(OUT)
print("ok", OUT)
