#!/usr/bin/env python3
"""Meldet URLs per IndexNow an Bing (und Yandex, Seznam, Naver).
Aufruf: python3 scripts/indexnow.py            -> alle URLs aus sitemap.xml
        python3 scripts/indexnow.py a.html b.html -> nur diese Seiten
Der Schlüssel liegt als <key>.txt im Repo-Wurzelverzeichnis (öffentlich, so verlangt es IndexNow)."""
import glob, json, os, re, sys, urllib.request
key_file = [f for f in glob.glob("*.txt") if re.fullmatch(r"[0-9a-f]{32}\.txt", f)][0]
key = key_file[:-4]
host = open("CNAME").read().strip()
base = f"https://{host}"
if len(sys.argv) > 1:
    urls = [f"{base}/{os.path.basename(a)}" for a in sys.argv[1:]]
else:
    urls = re.findall(r"<loc>([^<]+)</loc>", open("sitemap.xml", encoding="utf-8").read())
body = json.dumps({"host": host, "key": key, "keyLocation": f"{base}/{key_file}", "urlList": urls[:10000]}).encode()
req = urllib.request.Request("https://api.indexnow.org/indexnow", data=body, headers={"Content-Type": "application/json; charset=utf-8"})
try:
    with urllib.request.urlopen(req, timeout=30) as r:
        print(f"IndexNow: {len(urls)} URLs gemeldet, Antwort {r.status}")
except urllib.error.HTTPError as e:
    print(f"IndexNow: Fehler {e.code} {e.read()[:200]}")
