# rundfunkkritik.de — Dokumente & Werkzeuge

Dieses Repository enthält alle erstellten Dokumente, Skripte und Visualisierungen für **[rundfunkkritik.de](https://www.rundfunkkritik.de)**. Stand: Mai 2026.

---

## Inhalt

```
├── dokumente/          Dokumente als PDF (Website) und DOCX (Bearbeitung)
│   ├── *.pdf           → Download auf der Website, browserkompatibel
│   └── *.docx          → Zum Weiterbearbeiten / Anpassen
├── scripts/            Node.js-Skripte zum Regenerieren der Dokumente
└── visualisierung/     HTML-Visualisierungen — direkt einbettbar auf der Website
```

---

## Dokumente (`/dokumente`)

### `Klageschrift_Rundfunk_Bausteine_Mai2026.docx`
Modulare Klageschrift gegen den Rundfunkbeitrag — aufgeteilt in 8 eigenständige Bausteine statt 230 Seiten am Stück. Gerichte bevorzugen präzise, fokussierte Schriftsätze.

| Baustein | Inhalt | Wann einsetzen |
|----------|--------|----------------|
| A | Programmrüge — BVerwG-Linie + Vorgutachten | Hauptklage |
| B | Vollstreckungsangriff — BGH-Linie + Checkliste | Bei laufender Vollstreckung |
| C | Zeugen & Whistleblower (Herles, Tilgner, Schatz) | Beweisanträge Musterverfahren |
| D | DSGVO-Verstöße — separater Klageweg | Datenschutzrüge |
| E | Art. 3/5/19 GG — BVerfG-Vorlage | Musterverfahren / Revision |
| F | Sekundäre Darlegungslast — Prozessstrategie | Alle Verfahren |
| G | 22+ Gerichte mit Aktenzeichen (Tabelle) | Recherche & Verlinkung |
| H | 10 Vollstreckungsbehörden (Tabelle) | Vollstreckungsstrang |

> **Hinweis:** Kein Rechtsanwalt. Kein Ersatz für anwaltliche Beratung.

---

### `Fallbericht_5_Jahre_Widerstand_rundfunkkritik.docx`
Anonymisierter Erfahrungsbericht eines Betroffenen aus Bayern — 5 Jahre Widerstand gegen den Rundfunkbeitrag, ohne Anwalt. Chronologisch strukturiert mit konkreten Handlungsempfehlungen.

**Was der Bericht zeigt:**
- Wie § 766 ZPO eine laufende Vollstreckung stoppt
- Wie DSGVO-Auskunft Widersprüche aufdeckt
- Warum Zahlungen unter Vorbehalt rückforderbar sind
- Warum kompakte Klagen mehr bewirken als 200-Seiten-Vorlagen

---

### `Macht_Geld_OeRR_rundfunkkritik.docx`
Seriöse Analyse: Wer verdient was im ÖRR — und wer trifft täglich Entscheidungen über Millionen Menschen, ohne demokratisch gewählt zu sein.

**Kapitel:**
1. Politische Ebene: Gehälter & Pensionen (offiziell)
2. Intendanten: Strukturelle Macht, hohe Bezüge
3. Unsichtbare Entscheider: Programmdirektoren, Chefredakteure, CvD
4. Moderatoren: ZDF (bestätigte Zahlen) + ARD (Schätzungen)
5. Lebenszeitvolumen über Jahrzehnte
6. Vier Reformhebel

---

## Visualisierung (`/visualisierung`)

### `macht_geld_oerr.html`
Interaktive HTML-Seite mit vier Tabs:
- **Jahresgehalt** — horizontale Balken (ZDF bestätigt / ARD Schätzung / Politik)
- **Pension/Jahr** — Vergleich ÖRR vs. politische Ämter
- **Lebenszeit-Volumen** — gestapelt: Einkommen + Pension über die gesamte Karriere
- **Machtpyramide** — SVG-Visualisierung: 5 Ebenen, Breite = Gesellschaftsreichweite

**Einbettung auf der Website:**
```html
<iframe src="macht_geld_oerr.html" width="100%" height="700" frameborder="0"></iframe>
```
Oder direkt als HTML-Seite verlinken.

**Abhängigkeiten:** Chart.js 4.4.1 via CDN (cdnjs.cloudflare.com) — kein Build-Schritt nötig.

**Dark Mode:** Automatisch, folgt dem Systemthema.

---

## Skripte (`/scripts`)

Die Skripte regenerieren die Dokumente, falls Inhalte geändert werden sollen.

### Voraussetzungen
```bash
node -v    # Node.js 18+ empfohlen
npm install -g docx
```

### `klageschrift_bausteine.js`
Generiert `Klageschrift_Rundfunk_Bausteine_Mai2026.docx`.
```bash
node scripts/klageschrift_bausteine.js
# → Ausgabe: /mnt/user-data/outputs/Klageschrift_Rundfunk_Bausteine_Mai2026.docx
```

### `inhalte_rundfunkkritik.js`
Generiert beide Inhaltsdokumente (Fallbericht + Macht & Geld).
```bash
node scripts/inhalte_rundfunkkritik.js
# → Ausgabe: Fallbericht_5_Jahre_Widerstand_rundfunkkritik.docx
#           Macht_Geld_OeRR_rundfunkkritik.docx
```

> **Ausgabepfad anpassen:** In den Skripten `fs.writeFileSync(...)` auf einen lokalen Pfad ändern, z. B. `./dokumente/`.

---

## Rechtliche Hinweise

- Alle Dokumente: **kein Rechtsanwalt**, kein Ersatz für anwaltliche Beratung.
- ZDF-Gehaltsdaten: geleakte Liste (Welt am Sonntag, 2024) — bestätigt.
- ARD-Gehaltsdaten: Schätzungen auf Basis von Branchenquellen — als Schätzung gekennzeichnet.
- Fallbericht: vollständig anonymisiert, alle personenbezogenen Daten entfernt.
- Gerichtsurteile / Aktenzeichen: reale Verfahren, Stand Mai 2026.

---

## Lizenz

Inhalte zur freien Verwendung auf rundfunkkritik.de.
Für Weiterverwendung: Quellenangabe **rundfunkkritik.de** erwünscht.
