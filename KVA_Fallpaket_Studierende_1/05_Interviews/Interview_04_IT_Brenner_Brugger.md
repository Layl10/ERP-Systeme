# Interviewprotokoll 04 – IT
**Interviewpartner:** Markus Brenner (Leitung IT), Klaus Brugger (Entwicklung KVA-WARE, ab Min. 40) · **Datum:** 16.01.2026 · **Dauer:** 110 min · **Interviewerin:** L. Hartmann

---

## Teil 1 – Markus Brenner

**Architekturprinzipien (Vorschlag IT)**
- Standard-ERP als „Single Source of Truth" für Stammdaten (Artikel, Kunden, Lieferanten).
- PDM bleibt führend für Konstruktionsstücklisten → bidirektionale Schnittstelle Pflicht.
- Offene APIs (REST/OData), Event-basiert wünschenswert.
- Identity: Anbindung an Microsoft Entra ID (SSO), MFA.
- Keine neuen Access-/Excel-Lösungen.

**Betriebsmodell**
- Präferenz: Private Cloud / Hosting in Deutschland, oder On-Prem auf neuer Hardware.
- Argumente: Ausfallsicherheit Fertigung (Internet), Wehrtechnik, Datensouveränität, US CLOUD Act.
- „Ich bin nicht gegen Cloud. Ich bin gegen Cloud ohne Plan B."
- Akzeptabel wäre SaaS, wenn: EU-Anbieter oder EU-Datenhaltung mit vertraglicher Absicherung, redundante Internetanbindung (2. Leitung – ca. 18 T€/Jahr), Offline-Puffer für Rückmeldungen.

**Nutzerzahlen**
- „180 Named User im Projektauftrag sind gewürfelt." Seine Zählung: ca. **140 gleichzeitige** Nutzer max., aber ca. **230 Personen**, die irgendwann etwas sehen/buchen müssen (inkl. Lager, Meister, Servicetechniker, Pilsen).
- Zusätzlich: Shopfloor-Nutzer für Rückmeldungen (~ 210 Mitarbeitende) – Lizenzmodell kritisch!

**Ressourcen**
- IT kann max. 1,5 FTE fürs Projekt stellen. Ohne zusätzlichen Personal-/Partnereinsatz nicht machbar.
- Pilsen parallel: „Unmöglich."

**Termin**
- 01.01.2028 hält er nur für machbar, wenn Vertrag bis 09/2026 steht und Stammdatenbereinigung **jetzt** startet. Realistisch: Mitte 2028.

**Datenmigration**
- Artikel: ca. 52.000 aktiv, ca. 118.000 inkl. inaktiv/gesperrt.
- Stücklisten: ca. 38.000, Qualität schlecht.
- Offene Posten, offene Aufträge, Bestände: Stichtagsmigration.
- Historische Daten: GF will „10 Jahre Historie im neuen System" (Vertrieb), IT schlägt Archivlösung vor.

## Teil 2 – Klaus Brugger

- Bestätigt Ruhestand 31.12.2027. „Danach bin ich auf dem Boot am Bodensee."
- Auf die Frage nach Freelancer-Option: „Wenn es brennt, rufen Sie an. Aber ich mache keinen Vertrag." (→ Interpretationsspielraum!)
- Dokumentation KVA-WARE: „Gibt es teilweise. Das Wichtigste ist im Kopf."
- 340 Listen: „Davon werden vielleicht 40 regelmäßig genutzt. Ich sehe ja, was aufgerufen wird." Bietet an, eine Nutzungsstatistik zu erstellen.
- Hinweis: In KVA-WARE gibt es „versteckte" Logik, z. B. automatische Preisaufschläge für Kunde X, Sonderrabatte nach Region, Kupferzuschlag-Berechnung. „Das weiß niemand außer mir."
- Seine Meinung zum neuen System: „Nehmen Sie, was die Produktion versteht. Die Buchhaltung kommt mit allem klar."
