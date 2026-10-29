# Interviewprotokoll 06 – Produktion & Arbeitsvorbereitung
**Interviewpartner:** Stefan Wiedmann (Produktionsleitung), Petra Lang (AV / Fertigungssteuerung) · **Datum:** 22.01.2026 · **Dauer:** 120 min, inkl. Hallenrundgang · **Interviewerin:** L. Hartmann

---

## Grundhaltung
- Wiedmann: „Wir brauchen erst mal saubere Grundlagen: richtige Stücklisten, richtige Arbeitspläne, Material, das da ist. Dann reden wir über Feinplanung."
- Lang: „Ohne Feinplanung mit Kapazitätsgrenzen bringt uns das neue System gar nichts. Wir haben dann nur ein teureres KVA-WARE."
- → Beide einig: Härterei muss als Engpass abgebildet werden. Uneinig: APS als Muss in Phase 1?

## Anforderungen (gemeinsam erarbeitet)

**Stammdaten**
- Variantenstückliste/regelbasiert aus Konfigurator → automatisch Fertigungsstückliste + Arbeitsplan.
- PDM-Integration bidirektional mit Revisionsständen / Änderungswesen (ECM).
- Alternative Arbeitspläne (Singen/Pilsen, intern/extern).

**Planung**
- MRP mit realistischen Wiederbeschaffungszeiten, Losgrößenverfahren, Sicherheitsbeständen.
- Kapazitätsplanung (grob) für alle Bereiche.
- Feinplanung (finite) mindestens für Härterei und Verzahnung. Ofenbelegung: Chargenbildung nach Werkstoff + Härteverfahren + Ofengröße.
- Grafischer Leitstand (Gantt), Drag & Drop, Simulation von Eilaufträgen.
- Mehrwerksplanung Singen ↔ Pilsen (Bedarf, Kapazität, Transportzeit).

**Fertigung**
- Rückmeldung mobil (Tablet/Scanner), mit Mengen, Ausschuss, Störgründen.
- Ersatz BDE-Terminals (2012, Ersatzteile knapp!) – „bis spätestens 2027, egal wie".
- Fremdbearbeitung mit Beistellung, Lieferschein, Rückmeldung.
- Digitale Laufkarte / Arbeitspapiere mit Zeichnung.
- Werkzeug- und Vorrichtungsverwaltung (Wiedmann: Muss; Lang: „Kann").

**Instandhaltung**
- Wiedmann: Wartungspläne für 45 Maschinen, Störmeldungen, Ersatzteile Maschinen – „gehört ins ERP".
- Hinweis: GF Vogt hat das per Mail als „später" klassifiziert.

**Oberfläche**
- „F12-Ansicht": Liste aller Aufträge eines Bereichs mit Ampel, Material verfügbar ja/nein, nächster Arbeitsgang. Meister wollen „ein Bildschirm, alles drauf".
- Lang ergänzt: „Es muss nicht gleich aussehen, es muss gleich schnell sein."

**MES**
- Unklar, ob separates MES nötig. Wiedmann tendiert zu „ERP mit guter Fertigungskomponente", Lang zu „spezialisiertem MES/APS + ERP".
- Maschinendatenerfassung (MDE): Wunsch für Verzahnung (OPC UA verfügbar bei 6 Maschinen).

## Zahlen (Ergänzung zum Prozessdokument)
- Wiedmann nennt 14.500 FA p. a. Singen (Prozessdoku: 16.000) – „kommt drauf an, ob man Nacharbeit mitzählt".
- 6 Meister, 4 Personen AV, 3 Schichtführer Härterei.

## Betriebsrat-Bezug
- Rückmeldung personenbezogen nötig für Nachkalkulation? Lang: „Nein, Arbeitsplatz reicht." Controlling sieht das anders (Lohngruppen-Kosten).
