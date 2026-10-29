# Interviewprotokoll 09 – Service & After Sales
**Interviewpartner:** Andreas Seitz · **Datum:** 28.01.2026 · **Dauer:** 75 min · **Interviewerin:** L. Hartmann

---

## Ausgangslage
- Service ist laut Strategie 2030 der Wachstumstreiber (Ziel: Serviceumsatz verdoppeln, Wartungsverträge 41 → 300).
- „Mit heutiger IT ist das ausgeschlossen. Wir wissen nicht mal, welche Getriebe beim Kunden stehen."

## Anforderungen
**Installed Base / Seriennummern** – MUSS
- Jeder ausgelieferte Antrieb mit Seriennummer, Stückliste „as built", Kunde, Standort, Endkunde (falls Maschinenbauer weiterverkauft).
- Auch rückwirkend: Daten aus KVA-WARE seit ca. 2005 übernehmen (Fertigungsnummer = Pseudo-Seriennummer).
- „Ohne Seriennummer kein Wartungsvertrag, kein Ersatzteilshop, kein Predictive Maintenance."

**Serviceaufträge**
- Reparatur im Werk: Befundung, Kostenvoranschlag, Freigabe, Abrechnung.
- Außendienst: Einsatzplanung mit Qualifikationen, mobile App (Berichte, Fotos, Unterschrift, Zeiten, Material), **offline**, da Einsatzorte oft ohne Netz (Kranbahnen, Tunnel, Kellerräume).
- Wartungsverträge mit Intervallen, automatischer Auftragsauslösung, Vertragsabrechnung.
- Garantie/Kulanz mit Fehlercodes → Auswertung für Konstruktion & QM.

**Kundenportal / Ersatzteilshop**
- Kunde scannt QR-Code am Typenschild → sieht Ersatzteilstückliste seines Antriebs, bestellt online.
- Integriert mit Lagerbestand und Preisen aus ERP.
- (Vertrieb möchte auch Portal – siehe Interview 05. Abstimmung ein Portal für alle?)

**IoT / Zustandsüberwachung**
- Pilotprojekt mit Kranhersteller: Sensorik an 20 Getrieben. Daten liegen in einer Azure-Cloud des Kunden.
- „Das ERP muss das nicht können, aber es muss sich anbinden lassen."

## Systemfrage
- Seitz bevorzugt „ein spezialisiertes Field-Service-Tool, wenn das ERP es nicht gut kann" – kennt aus vorherigem Job eine Lösung. Gleichzeitig: „Nicht noch eine Insel."

## Organisation
- 6 Servicetechniker DE, 2 in CH (über KVA Antriebe AG), 4 Innendienst Service, 3 Werkstatt Reparatur.
- Servicetechniker Schweiz arbeiten heute komplett separat (Bexio für Rechnungen).

## Zahlen
- Siehe Prozessdokument. Ergänzung: ca. 180.000 ausgelieferte Antriebe seit 1990, davon geschätzt 60.000 noch im Einsatz.
