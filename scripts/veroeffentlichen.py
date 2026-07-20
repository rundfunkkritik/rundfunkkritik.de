#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
veroeffentlichen.py — Artikel auf dunkelfeld.report / rundfunkkritik.de bringen.

Aufruf (aus dem jeweiligen Repo-Verzeichnis):
    python3 scripts/veroeffentlichen.py NEUE-DATEI.html [WEITERE.html ...]
    python3 scripts/veroeffentlichen.py --pruefen          # nur prüfen, nichts ändern

Was das Skript macht (pro Datei):
  1. PRÜFT:  vollständiges HTML, keine <cite>-Reste, keine dunkelfeld.info-Altlasten,
             keine Telefonnummern, interne Links zeigen auf existierende Dateien.
  2. LIEST den Artikel-Kopf (Meta-Tags in der Datei):
             <meta name="rubrik" content="cat-medien">
             <meta name="kachel-titel" content="Titel für die Themen-Kachel">
  3. ERGÄNZT den Standard-Footer (Impressum/Datenschutz), falls er fehlt.
  4. TRÄGT die Themen-Kachel ein und zählt die Rubrik + Gesamtzahl hoch.
  5. ERGÄNZT die sitemap.xml.
  6. git add + commit. (Push macht die Session danach: git push)

Das Portal wird automatisch erkannt (CNAME-Datei im Repo).
"""
import os, re, sys, subprocess, datetime

# ---------------------------------------------------------------- Konfiguration
SITES = {
    "dunkelfeld.report": {
        "url": "https://dunkelfeld.report",
        "themen": "themen.html",
        "card": ('<a class="card" href="{fn}" data-s="{ds} {base}">'
                 '<span class="ct">{titel}</span><span class="cf">{fn}</span></a>'),
        "card_anchor": r'(<section class="cat" id="{rubrik}".*?)(</div>\s*</section>)',
        "counter_rubrik": r'(id="{rubrik}"><h2 class="cat-h">[^<]*<span class="cat-n">)(\d+)(</span>)',
        "counter_gesamt": r'(\d+)( Beiträge)',
        "footer": """<!-- LEGAL-FOOTER -->
<div style="max-width:720px;margin:2rem auto;padding:1.2rem;text-align:center;font-family:'Courier New',monospace;font-size:0.7rem;color:#52525b;border-top:1px solid #2a2a2e;">
  <a href="impressum.html" style="color:#e8c547;text-decoration:none;">Impressum</a> · <a href="datenschutz.html" style="color:#e8c547;text-decoration:none;">Datenschutz</a>
</div>
""",
        "sitemap_eintrag": '  <url><loc>{url}/{fn}</loc><lastmod>{datum}</lastmod><changefreq>monthly</changefreq><priority>0.7</priority></url>\n',
        "rubriken": ["cat-geld", "cat-stiftungen", "cat-medien", "cat-kontrolle", "cat-gewalt",
                     "cat-staat", "cat-migration", "cat-kirche", "cat-kinder", "cat-gruene",
                     "cat-nahrung", "cat-geo", "cat-endzeit", "cat-technik", "cat-seiten"],
    },
    "www.rundfunkkritik.de": {
        "url": "https://www.rundfunkkritik.de",
        "themen": "themen.html",
        "card": '      <li class="art"><a href="/{fn}">{titel}</a></li>',
        "card_anchor": r'(<section class="cat"><h2>{rubrik}</h2>\s*<ul>)',
        "counter_rubrik": None,          # rundfunkkritik hat keine Rubrik-Zähler
        "counter_gesamt": r'(\d+)( Beiträge)',
        "footer": """<footer style="text-align:center;padding:32px 16px;margin-top:48px;border-top:1px solid #d4c9b0;font-family:Georgia,'Crimson Pro',serif;font-size:14px;color:#6b6b6b">
  <a href="/themen.html" style="color:inherit;text-decoration:none">Themen</a> &middot; <a href="/impressum.html" style="color:#6b6b6b;text-decoration:none">Impressum</a> &middot; <a href="/datenschutz.html" style="color:#6b6b6b;text-decoration:none">Datenschutz</a> &middot; <a href="mailto:kontakt@rundfunkkritik.de" style="color:#6b6b6b;text-decoration:none">Kontakt</a>
</footer>
""",
        "sitemap_eintrag": '  <url>\n    <loc>{url}/{fn}</loc>\n    <lastmod>{datum}</lastmod>\n    <changefreq>monthly</changefreq>\n    <priority>0.7</priority>\n  </url>\n',
        "rubriken": ["Geld &amp; Finanzen", "Recht &amp; Beitrag", "Struktur &amp; Meinungsmache",
                     "Faktenchecks &amp; Einzelfälle", "Aufsicht, Personen &amp; Formate",
                     "Kampagne &amp; Mitmachen"],
    },
}

GRUEN, ROT, GELB, ENDE = "\033[92m", "\033[91m", "\033[93m", "\033[0m"

def lies(p):
    return open(p, encoding="utf-8").read()

def schreibe(p, s):
    open(p, "w", encoding="utf-8").write(s)

def site_erkennen():
    if not os.path.exists("CNAME"):
        sys.exit(f"{ROT}Keine CNAME-Datei — bitte aus dem Repo-Verzeichnis aufrufen.{ENDE}")
    dom = lies("CNAME").strip()
    for k in SITES:
        if k in dom or dom in k:
            return SITES[k]
    sys.exit(f"{ROT}Unbekanntes Portal: {dom}{ENDE}")

def pruefe(fn, cfg):
    """Gibt Liste von Problemen zurück (leer = alles gut)."""
    probleme = []
    s = lies(fn)
    if "</html>" not in s[-400:]:
        probleme.append("Datei unvollständig (kein </html> am Ende)")
    if "<cite" in s:
        probleme.append(f"{s.count('<cite')} <cite>-Reste (mit --reparieren automatisch entfernbar)")
    if "dunkelfeld.info" in s:
        probleme.append("alte dunkelfeld.info-Beschriftung")
    if re.search(r"\+4[39][0-9 ]{6,}", s):
        probleme.append("Telefonnummer im Artikel")
    for link in set(re.findall(r'href="([a-z0-9_\-]+\.html)"', s)):
        if not os.path.exists(link) and link != os.path.basename(fn):
            probleme.append(f"Link ins Leere: {link}")
    return probleme

def meta(fn, name):
    m = re.search(r'<meta name="' + name + r'" content="([^"]+)"', lies(fn))
    return m.group(1) if m else None

def titel_aus_datei(fn):
    m = re.search(r"<title>([^<|—]+)", lies(fn))
    return m.group(1).strip() if m else os.path.basename(fn)

def veroeffentliche(fn, cfg, reparieren=True):
    s = lies(fn)
    # 1. Reparaturen
    if reparieren and "<cite" in s:
        s = re.sub(r"<cite[^>]*>(.*?)</cite>", r"\1", s, flags=re.S)
    # 2. Footer
    if "impressum" not in s.lower():
        i = s.rfind("</body>")
        s = s[:i] + cfg["footer"] + s[i:]
    schreibe(fn, s)
    # 3. Themen-Kachel
    rubrik = meta(fn, "rubrik")
    titel = meta(fn, "kachel-titel") or titel_aus_datei(fn)
    t = lies(cfg["themen"])
    if fn in t:
        print(f"  {GELB}Kachel existiert schon — übersprungen.{ENDE}")
    elif not rubrik:
        print(f"  {GELB}Kein <meta name=\"rubrik\"> im Artikel — Kachel bitte von Hand oder Meta-Tag ergänzen.")
        print(f"  Verfügbare Rubriken: {', '.join(cfg['rubriken'])}{ENDE}")
    else:
        ds = titel.lower()
        for a, b in [("ä","ae"),("ö","oe"),("ü","ue"),("ß","ss"),("„",""),("“",""),('"',"")]:
            ds = ds.replace(a, b)
        card = cfg["card"].format(fn=fn, ds=ds, base=fn[:-5], titel=titel)
        anchor = cfg["card_anchor"].replace("{rubrik}", re.escape(rubrik))
        m = re.search(anchor, t, re.S)
        if not m:
            print(f"  {ROT}Rubrik '{rubrik}' nicht gefunden!{ENDE}")
        else:
            t = t[:m.end(1)] + ("\n" if card.startswith(" ") else "") + card + t[m.end(1):]
            if cfg["counter_rubrik"]:
                cr = cfg["counter_rubrik"].replace("{rubrik}", re.escape(rubrik))
                t = re.sub(cr, lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3), t, count=1)
            t = re.sub(cfg["counter_gesamt"], lambda m: str(int(m.group(1)) + 1) + m.group(2), t, count=1)
            schreibe(cfg["themen"], t)
            print(f"  Kachel in {rubrik} eingetragen, Zähler erhöht.")
    # 4. Sitemap
    sm = lies("sitemap.xml")
    if fn not in sm:
        eintrag = cfg["sitemap_eintrag"].format(url=cfg["url"], fn=fn,
                    datum=datetime.date.today().isoformat())
        i = sm.rfind("</urlset>")
        schreibe("sitemap.xml", sm[:i] + eintrag + sm[i:])
        print("  Sitemap ergänzt.")

def main():
    argv = [a for a in sys.argv[1:] if not a.startswith("--")]
    nur_pruefen = "--pruefen" in sys.argv
    cfg = site_erkennen()
    if not argv:
        sys.exit("Aufruf: python3 scripts/veroeffentlichen.py DATEI.html [...] [--pruefen]")
    fehler = 0
    for fn in argv:
        fn = os.path.basename(fn)
        print(f"\n{fn}:")
        probleme = pruefe(fn, cfg)
        for p in probleme:
            farbe = GELB if "cite" in p else ROT
            print(f"  {farbe}⚠ {p}{ENDE}")
        harte = [p for p in probleme if "cite" not in p]
        if harte:
            fehler += 1
            print(f"  {ROT}→ NICHT veröffentlicht (erst beheben).{ENDE}")
            continue
        if nur_pruefen:
            print(f"  {GRUEN}✓ prüfbereit{ENDE}")
            continue
        veroeffentliche(fn, cfg)
        print(f"  {GRUEN}✓ fertig{ENDE}")
    if not nur_pruefen and fehler == 0:
        subprocess.run(["git", "add", "-A"])
        subprocess.run(["git", "commit", "-m",
                        f"Neue Artikel: {', '.join(os.path.basename(a) for a in argv)}"])
        print(f"\n{GRUEN}Committet. Jetzt noch: git push{ENDE}")

if __name__ == "__main__":
    main()
