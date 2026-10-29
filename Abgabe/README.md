# Abgabe Fallstudie KVA – ERP 2028

| Artefakt | Datei |
|---|---|
| A – Management-Präsentation (5 Folien + Anhang A1–A12) | `KVA_ERP2028_Management-Praesentation.pptx` / `.pdf` |
| B + C – Anforderungskatalog, Widerspruchsliste, Nutzwertanalyse, TCO | `KVA_ERP2028_Anforderungen_Bewertung.xlsx` |
| D – KI-Nutzungsprotokoll (1 Seite) | `KVA_ERP2028_KI-Nutzungsprotokoll.docx` / `.pdf` |

Die Skripte, mit denen die Dateien erzeugt wurden, liegen in `skripte/` (`data.py` enthält alle Inhalte).

## Kernergebnisse (Version 2: zusammengeführt mit der ChatGPT-Gruppenversion)

- **42 Widersprüche** mit Fundstelle, **20 Störinformationen** bewusst aussortiert.
- **111 Rohanforderungen + 8 neue**, daraus **8 K.o.-Kriterien**.
- **Longlist mit 16 Optionen, Shortlist mit 4:** Business Central + Branchenlösung, proALPHA, IFS Cloud, abas ERP. APplus ist Nachrücker, weil für APplus kein Partner in Tschechien belegt ist.

| | Nutzwert | TCO 5 J. (Spanne) | Nutzwert je Mio. € |
|---|---|---|---|
| Business Central | 3,69 | 2,6 (2,2–3,5) Mio. € | **1,44** |
| proALPHA | 3,87 | 3,2 (2,7–4,3) Mio. € | 1,22 |
| IFS Cloud | **4,16** | 3,8 (3,2–5,1) Mio. € | 1,09 |
| abas | 3,64 | 2,9 (2,5–3,9) Mio. € | 1,25 |

- **Empfehlung:** Business Central als Favorit, aber erst nach einem Konfigurator-Proof-of-Concept. Rückfalloption ist proALPHA. Pilsen geht als Pilot zuerst live (10/2027), Singen am 01.04.2028.
- **Budget:** Keine Option bleibt unter 1,2 bzw. 1,5 Mio. €. Die E-Rechnungs-Pflicht gilt für KVA schon ab dem **01.01.2027**.

## Was aus der Gruppenversion übernommen wurde

- **Korrigierte Partner in Tschechien**, nach Prüfung der Quellen: proALPHA (SPC solutions, Zlín), abas (amotIQ), Infor (ITeuro).
- **Kostenspannen** niedrig/Basis/hoch.
- **Gemeinsames Testskript** T1–T6 für alle Anbieter (Anhang A11, Excel-Blatt 15).
- **Nummerierte Quellen** S01–S23 (Anhang A12, Excel-Blatt 10).
- **Business-Central-Preise** in Euro von der Microsoft-Preisseite.

## Was aus den Vorlesungsfolien übernommen wurde

- **Neues Kriterium „Systemarchitektur Multi-Site vs. Mehrmandanten“** (Kurz 2026; Martin/Kurz 2024). Es geht darum, ob Intercompany-Aufträge Pilsen → Singen intern laufen oder über Belegaustausch.
- **Keine Marktanteilsstudien**, denn laut Vorlesung sind die „Datenquellen meist zweifelhaft“.
- **Länderversionen:** für die Schweiz z. B. die 5-Rappen-Rundung.

## Umsetzung der Dozenten-Anmerkungen

| Anmerkung | Wo umgesetzt |
|---|---|
| Name + Matrikelnummer auf Folie 1 | dunkles Feld unter dem Titel (Platzhalter) |
| Gescheitertes ERP-Projekt 2016 | Folie 1, Folie 5 (Bedingung), Anhang A5 |
| Störinformationen gegen KI | Anhang A9, Excel-Blatt 11 |
| Passende ERP-Systeme, Lizenzen | Folie 3, Anhang A7 + A8, Excel-Blätter 5 + 12 |
| Gesamtkosten hochrechnen | Folie 4, Anhang A4, Excel-Blätter 7–9 |
| Nächste Schritte, Ausschreibung, Risiken | Folie 5, Anhang A5, A6, A11 |
| Nutzwertanalyse | Folie 4, Anhang A3, Excel-Blatt 6 (mit Sensitivität) |
| HTWG-Design | Farben, Wortmarke, Seitenzahl und Tabellenstil wie die Vorlesungsfolien |
| Management „abholen“ | Folie 5 (Tabelle je Person), Sprechernotizen, Excel-Blatt 14 |
| EU-Ausland / Budget | Folie 2, Anhang A10, Excel-Blatt 13; Budget auf Folie 4 und 5 |

## Vor der Abgabe vom Team zu erledigen

1. Namen und Matrikelnummern auf Folie 1 und im KI-Protokoll eintragen.
2. Punkte der Nutzwertanalyse und TCO-Annahmen (gelbe Zellen) im Team prüfen.
3. Quellen, die als Herstellerangabe oder „Abruf durch Projektgruppe“ markiert sind, selbst im Browser öffnen: S06, S08, S18.
