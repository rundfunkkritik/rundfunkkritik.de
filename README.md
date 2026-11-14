# rundfunkkritik.de

Quellcode der Website **[rundfunkkritik.de](https://www.rundfunkkritik.de)**: strukturelle Medienanalyse zum öffentlich-rechtlichen Rundfunk. Statisches HTML, ausgeliefert über GitHub Pages (Domain in `CNAME`). Stand: November 2026.

## Aufbau

| Pfad | Inhalt |
|---|---|
| `index.html` | Startseite mit den aktuellen Analysen als Karten |
| `themen.html` | Alle Beiträge nach Rubriken, mit Gesamtzähler |
| `*.html` | Einzelne Artikel (flach im Hauptverzeichnis) |
| `impressum.html`, `datenschutz.html` | Rechtliches |
| `*.pdf` | Downloads (Klageschrift-Bausteine, Fallbericht, Beweisdossier, Macht & Geld) |
| `mail.js` | Kontakt-Skript, wird von fast allen Seiten geladen – nicht löschen |
| `search-index.json` | Suchindex der Startseiten-Suche |
| `sitemap.xml`, `robots.txt` | Für Suchmaschinen |
| `fonts/`, `fonts.css`, `vendor/` | Selbst gehostete Schriften und Chart.js (DSGVO) |
| `scripts/` | Werkzeuge (siehe unten) |
| `klageschrift_bausteine.js`, `inhalte_rundfunkkritik.js`, `package.json` | Node-Generatoren für die Word/PDF-Dokumente |

## Neuen Artikel veröffentlichen

```bash
python3 scripts/veroeffentlichen.py NEUER-ARTIKEL.html
```

Das Skript prüft den Artikel (vollständiges HTML, interne Links), ergänzt den Footer, trägt ihn in `themen.html` ein, zählt hoch, ergänzt die `sitemap.xml` und committet. Danach `git push`. Für eine Karte auf der Startseite einen `feature-card`-Block oben in `index.html` einfügen.

Weitere Werkzeuge:

- `python3 scripts/build_search_index.py` – Suchindex neu erzeugen
- `python3 scripts/indexnow.py [datei.html …]` – geänderte Seiten an Bing/IndexNow melden

## Regeln

- Nur veröffentlichungsfähige Dateien ins Repo: Alles hier ist öffentlich abrufbar, auch ohne Verlinkung.
- Arbeitsdateien, Entwürfe und Word-Fassungen gehören nicht ins Repo.
- Interne Links immer auf existierende Dateinamen setzen (`/datei.html`).
