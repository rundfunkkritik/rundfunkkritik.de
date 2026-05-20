const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  ExternalHyperlink, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

// ─── FARBEN & STILE ──────────────────────────────────────────────────────────
const C = {
  dark:      "1A2F4A",
  mid:       "2E5C8A",
  light:     "D0E4F5",
  accent:    "C0392B",
  gold:      "8B6914",
  grey:      "888888",
  lightgrey: "F2F5F8",
  white:     "FFFFFF"
};
const bn = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
const bt = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const bm = { style: BorderStyle.SINGLE, size: 4, color: C.mid   };
const ba = { style: BorderStyle.SINGLE, size: 4, color: C.accent };
const allNone = { top: bn, bottom: bn, left: bn, right: bn };
const allThin = { top: bt, bottom: bt, left: bt, right: bt };

// ─── HELPER ──────────────────────────────────────────────────────────────────
const r = (text, o={}) => new TextRun({
  text, font:"Arial", size: o.size||22,
  color: o.color||"333333", bold:!!o.bold, italics:!!o.italic
});
const rw = (text, o={}) => new TextRun({
  text, font:"Arial", size: o.size||22, color: C.white, bold:!!o.bold
});

const p = (text, o={}) => new Paragraph({
  spacing:{ before: o.before||80, after: o.after||120 },
  alignment: o.align||AlignmentType.LEFT,
  children:[ r(text, o) ]
});

const pR = (children, o={}) => new Paragraph({
  spacing:{ before: o.before||80, after: o.after||120 },
  alignment: o.align||AlignmentType.LEFT,
  children
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing:{ before:480, after:160 },
  children:[r(text,{size:38,bold:true,color:C.dark})]
});
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing:{ before:280, after:100 },
  children:[r(text,{size:26,bold:true,color:C.mid})]
});
const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing:{ before:200, after:80 },
  children:[r(text,{size:22,bold:true,color:C.dark})]
});

const bull = (text, o={}) => new Paragraph({
  numbering:{ reference:"bullets", level:0 },
  spacing:{ before:60, after:60 },
  children:[r(text,o)]
});

const rule = (color=C.mid) => new Paragraph({
  spacing:{ before:200, after:200 },
  border:{ bottom:{ style:BorderStyle.SINGLE, size:6, color, space:1 } },
  children:[]
});

const space = (pt=120) => new Paragraph({
  spacing:{ before:pt, after:0 }, children:[r("")]
});

const pb = () => new Paragraph({ children:[new TextRun({ break:1 })] });

const link = (text, url) => new ExternalHyperlink({
  link:url,
  children:[new TextRun({ text, font:"Arial", size:22, color:C.mid, underline:{} })]
});

// Box mit Hintergrundfarbe
const infoBox = (label, lines, bg=C.lightgrey) => new Table({
  width:{ size:9026, type:WidthType.DXA },
  columnWidths:[9026],
  rows:[new TableRow({
    children:[new TableCell({
      borders:allNone,
      shading:{ fill:bg, type:ShadingType.CLEAR },
      margins:{ top:160, bottom:160, left:240, right:240 },
      width:{ size:9026, type:WidthType.DXA },
      children:[
        new Paragraph({ spacing:{before:0,after:100},
          children:[r(label,{bold:true,color:C.mid,size:20})] }),
        ...lines.map(l => new Paragraph({ spacing:{before:0,after:60},
          children:[r(l,{size:20,color:"444444"})] }))
      ]
    })]
  })]
});

// Zitat-Box
const zitat = (text, quelle) => new Table({
  width:{ size:9026, type:WidthType.DXA },
  columnWidths:[200, 8826],
  rows:[new TableRow({
    children:[
      new TableCell({
        borders:allNone,
        shading:{ fill:C.mid, type:ShadingType.CLEAR },
        width:{ size:200, type:WidthType.DXA },
        children:[p("")]
      }),
      new TableCell({
        borders:allNone,
        shading:{ fill:C.lightgrey, type:ShadingType.CLEAR },
        margins:{ top:140, bottom:140, left:240, right:240 },
        width:{ size:8826, type:WidthType.DXA },
        children:[
          new Paragraph({ spacing:{before:0,after:80},
            children:[r(`'${text}"`,{italic:true,color:"333333",size:23})] }),
          new Paragraph({ spacing:{before:0,after:0},
            children:[r(quelle,{color:C.grey,size:18})] })
        ]
      })
    ]
  })]
});

// Tabellen-Zeile zweifarbig
const tRow = (cells, widths, isHeader=false, shade="#FFFFFF") => new TableRow({
  tableHeader: isHeader,
  children: cells.map((text,i) => new TableCell({
    borders: allThin,
    shading:{ fill: isHeader ? C.dark : shade, type:ShadingType.CLEAR },
    margins:{ top:80, bottom:80, left:140, right:140 },
    width:{ size:widths[i], type:WidthType.DXA },
    children:[new Paragraph({ children:[
      isHeader
        ? rw(text,{size:19,bold:true})
        : r(text,{size:19,color:"333333"})
    ]})]
  }))
});

// ═══════════════════════════════════════════════════════════════════════════════
// DOKUMENT 1 — FALLBERICHT
// ═══════════════════════════════════════════════════════════════════════════════

const fallbericht = [

  // Titel
  new Paragraph({ spacing:{before:800,after:80},
    children:[r("Fünf Jahre Widerstand",{size:56,bold:true,color:C.dark})] }),
  new Paragraph({ spacing:{before:0,after:60},
    children:[r("Ein anonymisierter Erfahrungsbericht aus Bayern",{size:28,color:C.mid})] }),
  new Paragraph({ spacing:{before:0,after:60},
    children:[r("Für rundfunkkritik.de  ·  Stand Mai 2026  ·  Zur freien Verwendung",{size:20,italic:true,color:C.grey})] }),
  rule(C.mid),
  infoBox("Warum dieser Bericht wichtig ist",[
    "Dieser Fall ist kein Einzelfall. Er zeigt, wie Beitragsverfahren in der Praxis ablaufen —",
    "von der ersten Zahlungsverweigerung bis zur laufenden Vollstreckung, von der Stille",
    "der Behörden bis zur wirksamen Gegenstrategie. Alle personenbezogenen Daten wurden",
    "anonymisiert. Der Betroffene hat das Verfahren ohne Anwalt und ohne digitale Hilfe",
    "in der ersten Phase durchgefochten. Seine Erfahrungen sind dokumentiert."
  ]),
  space(200),

  // Phase 1
  h1("2021 — Der Beginn: Widerspruch ohne Anwalt"),
  p("Im Sommer 2021 widerrief ein Bürger aus Südbayern sein SEPA-Lastschriftmandat. Es folgte eine Rücklastschrift, dann ein Festsetzungsbescheid des Beitragsservice. Der Betroffene legte fristgerecht Widerspruch ein — handgeschrieben, ohne juristische Unterstützung, mit dem Argument strukturell unausgewogener Berichterstattung."),
  p("Dann: Stille. Vier Jahre lang."),
  space(),

  zitat("Der Widerspruch war handgeschrieben, ohne Vorlage, ohne Anwalt. Einfach präzise beschrieben, was nicht stimmt.","Aus dem anonymisierten Bericht"),
  space(),

  h2("Was man daraus lernen kann"),
  bull("Widerspruch immer per Einschreiben mit Rückschein — das ist der Zustellungsnachweis."),
  bull("Programmkritik allein trägt nicht — sie ist kein anerkannter Befreiungsgrund."),
  bull("Schweigen der Behörde ist keine Zustimmung — aber es kann Verwirkungsargumente eröffnen."),
  space(200),

  // Phase 2
  h1("2021–2025 — Die Funkstille und was sie bedeutet"),
  p("Von 2021 bis Anfang 2025 kam kein einziges Schreiben. Ende 2023 zog der Betroffene um. Die neue Adresse wurde automatisch beim Zuzug an den Beitragsservice übermittelt — das belegte später eine DSGVO-Auskunft nach Art. 15 DSGVO eindeutig."),
  p("Trotz bekannter Adresse: weitere 21 Monate Schweigen."),
  space(),

  infoBox("Rechtlicher Hinweis: Verwirkung",[
    "Wer einen Anspruch über Jahre nicht geltend macht, obwohl er es hätte tun können,",
    "riskiert die Verwirkung. Das ist kein Freifahrtschein — aber ein prozessual relevantes Argument,",
    "das gezielt in eine Klage eingebracht werden kann."
  ]),
  space(),

  h2("Der DSGVO-Winkel: Was die Auskunft offenbarte"),
  p("Die Auskunft nach Art. 15 DSGVO brachte einen bemerkenswerten Widerspruch ans Licht: In der allgemeinen Datenschutzerklärung des Beitragsservice wurden Kategorien wie Telefongesprächsnotizen, archivierten Schriftwechsel und Mahnhistorie als verarbeitet aufgeführt. Die tatsächliche Auskunft enthielt keine einzige dieser Kategorien."),
  p("Zusätzlich: Die Gemeinde bestätigte schriftlich, dass keine automatisierten Abrufverfahren stattfanden. Seit Juni 2025 besteht sogar eine unbefristete Übermittlungssperre. Der Beitragsservice bezeichnete die Adressherkunft dennoch als 'anlassbezogen' — ein Widerspruch, den die Datenschutzbehörde nun prüft."),
  space(),

  zitat("Die Gemeinde sagte: keine automatisierten Abrufe. Der Beitragsservice sagte: anlassbezogen. Beides kann nicht gleichzeitig stimmen.","Aus dem anonymisierten Bericht"),
  space(200),
  pb(),

  // Phase 3
  h1("Februar 2025 — Der Widerspruchsbescheid kommt"),
  p("Dreieinhalb Jahre nach dem Widerspruch: Zustellung des Widerspruchsbescheids. Zurückweisung mit der Begründung, Programmkritik sei kein zulässiger Grund, die Beitragspflicht zu verweigern. Frist: ein Monat für Klage beim Verwaltungsgericht."),
  p("Der Betroffene erhob Klage — diesmal strukturiert, ohne Massenvorlage, mit drei konkreten Argumenten:"),
  bull("Verjährung der ältesten Forderungen"),
  bull("fehlende Beitragspflicht nach Umzug (Bruder zahlt bereits für dieselbe Wohnung)"),
  bull("fehlender Zustellungsnachweis für die zugrundeliegenden Bescheide"),
  space(),

  infoBox("Tipp: Kompakte Klage statt 200-Seiten-Block",[
    "Gerichte lesen keine überladenen Schriftsätze. Eine präzise, auf 3–5 konkrete Punkte",
    "reduzierte Klage wird gelesen und verarbeitet. 200-seitige Musterklagen werden als",
    "pauschal eingestuft — und entsprechend behandelt."
  ]),
  space(200),

  // Phase 4
  h1("Herbst 2025 — Die Vollstreckungsmaschine läuft an"),
  p("Noch während die Klage beim Verwaltungsgericht anhängig war, beauftragte der Beitragsservice die Gerichtsvollzieherin. Der Betroffene wurde zur Vermögensauskunft geladen. Er zahlte — aber ausschließlich unter ausdrücklichem Vorbehalt und Zwang, mit entsprechendem Vermerk auf jeder Überweisung."),
  p("Insgesamt flossen so über fünf Monate 510 Euro."),
  space(),

  zitat("Zahlung unter Vorbehalt und Zwang ist keine freiwillige Leistung. Sie ist kondiktionsfähig — rückforderbar, wenn die Forderung nicht bestätigt wird.","Rechtlicher Kontext § 812 BGB"),
  space(),

  h2("Was Zahlungen unter Vorbehalt bedeuten"),
  bull("Jede Zahlung muss explizit als 'unter Vorbehalt und Zwang' bezeichnet werden."),
  bull("Der Vorbehalt muss auf dem Überweisungsträger vermerkt sein."),
  bull("Er muss bei jeder einzelnen Zahlung wiederholt werden — nicht nur einmalig."),
  bull("Bei Obsiegen im Hauptverfahren: Rückforderungsklage nach § 812 BGB möglich."),
  space(200),

  // Phase 5
  h1("Januar 2026 — Die Erinnerung nach § 766 ZPO wirkt"),
  p("Der Betroffene stellte fest: Für die zugrundeliegenden Festsetzungsbescheide existiert kein ordnungsgemäßer Zustellungsnachweis. Kein Einschreiben-Rückschein, keine Zustellungsurkunde. Auch das Vollstreckungsersuchen selbst wies eine formunwirksame Signatur auf."),
  p("Er stellte eine Erinnerung nach § 766 ZPO. Im März 2026 kam die Antwort: Die Zwangsvollstreckung wurde vorläufig eingestellt."),
  space(),

  infoBox("§ 766 ZPO — Das wirksamste kurzfristige Mittel bei laufender Vollstreckung",[
    "Kostet nichts. Braucht keinen Anwalt. Kann schnell wirken.",
    "Voraussetzung: formelle Mängel des Vollstreckungsersuchens (fehlende Signatur,",
    "fehlender Zustellungsnachweis, fehlende Verantwortlichkeitszurechnung).",
    "BGH VII ZB 29/24 (Februar 2026) hat die Anforderungen deutlich geschärft."
  ]),
  space(),

  h2("Mai 2026 — Nächster Anlauf, nächste Reaktion"),
  p("Der Beitragsservice reagierte mit einem neuen Vollstreckungsauftrag — ohne die festgestellten Mängel zu beheben. Der Betroffene legte erneut Erinnerung nach § 766 ZPO ein, stellte gleichzeitig einen Antrag auf aufschiebende Wirkung nach § 80 Abs. 5 VwGO und reichte eine DSGVO-Beschwerde beim Bayerischen Landesbeauftragten für Datenschutz ein."),
  space(200),
  pb(),

  // Fazit
  h1("Was funktioniert — und was nicht"),
  space(80),

  h2("Wirksam"),
  bull("Erinnerung nach § 766 ZPO bei formellen Mängeln des Vollstreckungsersuchens"),
  bull("Zahlung unter ausdrücklichem Vorbehalt und Zwang — schriftlich bei jeder Zahlung"),
  bull("DSGVO-Auskunft nach Art. 15 DSGVO als Werkzeug zur Aufdeckung von Widersprüchen"),
  bull("Auskunft bei der Einwohnermeldebehörde nach § 10/11 BMG"),
  bull("Kompakte, präzise Klageschrift — auf konkrete Punkte reduziert"),
  bull("Alle Schreiben per Einschreiben mit Rückschein"),
  space(),

  h2("Weniger wirksam vor Gericht"),
  bull("Programmkritik und journalistische Qualitätsmängel allein als Klagegrund"),
  bull("Gewissensfreiheit nach Art. 4 GG — vom BVerfG nicht als Befreiungsgrund anerkannt"),
  bull("Überlange Schriftsätze — Gerichte behandeln 200-Seiten-Vorlagen als pauschal"),
  space(200),

  // Schluss
  rule(C.accent),
  new Paragraph({ spacing:{before:200,after:80}, alignment:AlignmentType.CENTER,
    children:[r("Stand: Mai 2026  ·  rundfunkkritik.de  ·  Alle personenbezogenen Daten anonymisiert",
      {size:18,italic:true,color:C.grey})] }),
  new Paragraph({ spacing:{before:0,after:0}, alignment:AlignmentType.CENTER,
    children:[r("Kein Rechtsanwalt — kein Ersatz für anwaltliche Beratung.",
      {size:18,italic:true,color:C.grey})] }),
];


// ═══════════════════════════════════════════════════════════════════════════════
// DOKUMENT 2 — MACHT & GELD IM ÖRR
// ═══════════════════════════════════════════════════════════════════════════════

const machtGeld = [

  // Titel
  new Paragraph({ spacing:{before:800,after:80},
    children:[r("Wer Deutschland wirklich prägt",{size:52,bold:true,color:C.dark})] }),
  new Paragraph({ spacing:{before:0,after:60},
    children:[r("Macht, Geld und Einfluss im öffentlich-rechtlichen Rundfunk",{size:28,color:C.mid})] }),
  new Paragraph({ spacing:{before:0,after:60},
    children:[r("Eine Analyse für rundfunkkritik.de  ·  Stand Mai 2026",{size:20,italic:true,color:C.grey})] }),
  rule(C.mid),
  infoBox("Zur Methodik",[
    "ZDF-Gehälter: geleakte Liste (Welt am Sonntag, 2024) — verifizierte Zahlen.",
    "ARD-Gehälter: keine offizielle Veröffentlichung, Schätzungen auf Basis von Branchenleaks,",
    "Medienberichten und Sendefrequenz — als Schätzung gekennzeichnet.",
    "Pensionen ÖRR: Kombination aus öffentlichen Angaben (RBB-Skandal) und Branchen-",
    "schätzungen. Politische Bezüge: offiziell nach Bundesbesoldungsgesetz."
  ]),
  space(200),

  // Kapitel 1 — Politik
  h1("I. Die politische Ebene: Formale Macht, transparente Bezüge"),
  p("Wer Macht hat, verdient viel — so lautet die verbreitete Annahme. Im deutschen politischen System stimmt das nur bedingt. Die Bezüge der politischen Spitze sind gesetzlich geregelt, öffentlich einsehbar und im internationalen Vergleich moderat."),
  space(),

  new Table({
    width:{ size:9026, type:WidthType.DXA },
    columnWidths:[3000,2200,2200,1626],
    rows:[
      tRow(["Funktion","Jahresgehalt","Jahrespension","Besonderheit"],[3000,2200,2200,1626],true),
      tRow(["Bundeskanzler","~ 360.000 €","bis 250.000 €","gesetzlich geregelt"],[3000,2200,2200,1626],false,"FAFBFC"),
      tRow(["Bundesminister","~ 280.000 €","180.000–220.000 €","transparent"],[3000,2200,2200,1626]),
      tRow(["Staatsekretär","~ 220.000 €","120.000–160.000 €","transparent"],[3000,2200,2200,1626],false,"FAFBFC"),
      tRow(["Bundestagsabgeordnete","~ 120.000 €","ca. 7.000 € / Monat","demokratisch legitimiert"],[3000,2200,2200,1626]),
    ]
  }),
  space(),
  p("Das entscheidende Merkmal: Jeder dieser Beträge ist öffentlich, gesetzlich geregelt und demokratisch legitimiert. Politiker werden gewählt, können abgewählt werden und stehen unter permanenter Öffentlichkeitskontrolle.", {color:C.grey,italic:true}),
  space(200),

  // Kapitel 2 — Intendanten
  h1("II. Die Intendanten: Strukturelle Macht, hohe Bezüge"),
  p("Intendanten leiten Institutionen, die täglich Millionen Menschen erreichen. Sie kontrollieren Milliarden-Budgets, entscheiden über Personal und prägen die strategische Ausrichtung ganzer Sender. Ihre Bezüge übersteigen die ministerialer Funktionen — bei deutlich geringerer demokratischer Kontrolle."),
  space(),

  new Table({
    width:{ size:9026, type:WidthType.DXA },
    columnWidths:[2800,2000,2000,2226],
    rows:[
      tRow(["Intendant / Anstalt","Jahresgehalt","Jahrespension (Schätzung)","Kontrollstruktur"],[2800,2000,2000,2226],true),
      tRow(["Tom Buhrow (WDR)","~ 413.000 €","120.000–180.000 €","Rundfunkrat (pol. besetzt)"],[2800,2000,2000,2226],false,"FAFBFC"),
      tRow(["Norbert Himmler (ZDF)","~ 372.000 €","100.000–160.000 €","Fernsehrat (pol. besetzt)"],[2800,2000,2000,2226]),
      tRow(["Kai Gniffke (SWR)","~ 361.000 €","100.000–150.000 €","Rundfunkrat"],[2800,2000,2000,2226],false,"FAFBFC"),
      tRow(["Katja Wildermuth (BR)","~ 340.000 €","100.000–150.000 €","Rundfunkrat"],[2800,2000,2000,2226]),
      tRow(["Joachim Knuth (NDR)","~ 346.000 €","100.000–150.000 €","Rundfunkrat"],[2800,2000,2000,2226],false,"FAFBFC"),
      tRow(["P. Schlesinger (RBB, †)","~ 303.000 € + Boni","~ 180.000 € / Jahr","nicht extern geprüft"],[2800,2000,2000,2226]),
    ]
  }),
  space(),

  zitat("Intendanten verdienen mehr als Minister — und werden nicht gewählt.","Analyse rundfunkkritik.de"),
  space(200),
  pb(),

  // Kapitel 3 — Unsichtbare Entscheider
  h1("III. Die unsichtbaren Entscheider: Programm-Architekten"),
  p("Sie entscheiden täglich, welche Themen laufen, welche Experten sprechen, welche Bilder gezeigt werden und welche Perspektiven fehlen. Ihr Name steht selten in der Zeitung. Ihre Macht ist operativ, täglich und direkt."),
  space(),

  new Table({
    width:{ size:9026, type:WidthType.DXA },
    columnWidths:[2800,2000,1900,2326],
    rows:[
      tRow(["Funktion","Jahresgehalt","Pension/Jahr","Machtbereich"],[2800,2000,1900,2326],true),
      tRow(["Programmdirektor","250.000–300.000 €","80.000–120.000 €","Strategische Sendepolitik"],[2800,2000,1900,2326],false,"FAFBFC"),
      tRow(["Chefredakteur","180.000–250.000 €","60.000–90.000 €","Nachrichtenagenda"],[2800,2000,1900,2326]),
      tRow(["Chef vom Dienst (CvD)","90.000–150.000 €","30.000–50.000 €","Tägliche Themenauswahl"],[2800,2000,1900,2326],false,"FAFBFC"),
      tRow(["Leitender Redakteur","70.000–120.000 €","25.000–40.000 €","Ressort-Agenda"],[2800,2000,1900,2326]),
    ]
  }),
  space(),
  p("Diese Gruppe ist kaum sichtbar — aber sie ist es, die entscheidet, welche Realität täglich vermittelt wird. Keine dieser Funktionen ist demokratisch gewählt. Keine unterliegt einer externen Transparenzpflicht.", {color:C.grey,italic:true}),
  space(200),

  // Kapitel 4 — Moderatoren
  h1("IV. Die Moderatoren: Sichtbare Macht, sehr hohe Honorare"),
  p("Moderatoren sind die Gesichter des Systems. Sie genießen Vertrauen, prägen Debatten und erreichen täglich Millionen. Viele sind nicht angestellt, sondern als freie Mitarbeiter tätig — mit entsprechend höheren Honoraren und eigenen Produktionsfirmen."),
  space(),

  p("ZDF — bestätigte Zahlen aus geleakter Liste (Welt am Sonntag, 2024):"),
  new Table({
    width:{ size:9026, type:WidthType.DXA },
    columnWidths:[2600,1800,4626],
    rows:[
      tRow(["Name","Honorar / Jahr","Format / Funktion"],[2600,1800,4626],true),
      tRow(["Markus Lanz","1.900.000 €","Talkshow, Dokus, Sondermoderation"],[2600,1800,4626],false,"FAFBFC"),
      tRow(["Horst Lichter","1.700.000 €","Bares für Rares"],[2600,1800,4626]),
      tRow(["Oliver Welke","1.180.000 €","heute-show"],[2600,1800,4626],false,"FAFBFC"),
      tRow(["Jan Böhmermann","682.000 €","ZDF Magazin Royale"],[2600,1800,4626]),
      tRow(["Johannes B. Kerner","630.000 €","Sport, Shows, Moderation"],[2600,1800,4626],false,"FAFBFC"),
      tRow(["Maybrit Illner","480.000 €","Politiktalk"],[2600,1800,4626]),
      tRow(["Andrea Kiewel","400.000 €","ZDF-Fernsehgarten"],[2600,1800,4626],false,"FAFBFC"),
      tRow(["Marietta Slomka","393.750 €","heute-journal"],[2600,1800,4626]),
      tRow(["Rudi Cerne","382.000 €","Aktenzeichen XY"],[2600,1800,4626],false,"FAFBFC"),
      tRow(["Christian Sievers","350.000 €","heute-journal"],[2600,1800,4626]),
    ]
  }),
  space(),

  p("ARD — keine offizielle Veröffentlichung, Schätzungen auf Basis von Branchenquellen:"),
  new Table({
    width:{ size:9026, type:WidthType.DXA },
    columnWidths:[2600,2200,4226],
    rows:[
      tRow(["Name","Honorar-Schätzung","Format"],[2600,2200,4226],true),
      tRow(["Kai Pflaume","1.000.000–1.500.000 €","Quizshows, Unterhaltung"],[2600,2200,4226],false,"FAFBFC"),
      tRow(["Anne Will","1.000.000–1.500.000 €","Talkshow (bis 2023)"],[2600,2200,4226]),
      tRow(["Sandra Maischberger","1.000.000–1.400.000 €","Talkshow"],[2600,2200,4226],false,"FAFBFC"),
      tRow(["Frank Plasberg","1.000.000–1.200.000 €","Hart aber fair (bis 2022)"],[2600,2200,4226]),
      tRow(["Eckart v. Hirschhausen","800.000–1.200.000 €","Gesundheitsformate"],[2600,2200,4226],false,"FAFBFC"),
      tRow(["Jörg Pilawa","600.000–900.000 €","Quizshows"],[2600,2200,4226]),
      tRow(["Carolin Kebekus","600.000–900.000 €","Comedy, Late Night"],[2600,2200,4226],false,"FAFBFC"),
    ]
  }),
  space(200),
  pb(),

  // Kapitel 5 — Lebenszeitvolumen
  h1("V. Die Dimension über Jahrzehnte"),
  p("Einzelne Jahresgehälter sind schwer vorstellbar. Was die Dimension wirklich deutlich macht, sind die Lebenszeitvolumen — also die Summe aus Einkommen und Pension über eine typische Karriere."),
  space(),

  new Table({
    width:{ size:9026, type:WidthType.DXA },
    columnWidths:[2600,1400,1600,1600,1826],
    rows:[
      tRow(["Gruppe","Karriere","Einkommen gesamt","Pension gesamt","Lebenszeitvolumen"],[2600,1400,1600,1600,1826],true),
      tRow(["Moderatoren","20 Jahre","20–30 Mio. €","privat, mehrere Mio.","25–40 Mio. €"],[2600,1400,1600,1600,1826],false,"FAFBFC"),
      tRow(["Intendanten","20 Jahre","7–8 Mio. €","3–4 Mio. €","10–12 Mio. €"],[2600,1400,1600,1600,1826]),
      tRow(["Programmdirektoren","25 Jahre","6–7 Mio. €","2–3 Mio. €","8–10 Mio. €"],[2600,1400,1600,1600,1826],false,"FAFBFC"),
      tRow(["Chefredakteure","25 Jahre","4–5 Mio. €","1,5–2 Mio. €","5–7 Mio. €"],[2600,1400,1600,1600,1826]),
      tRow(["Politiker (Kanzler)","15–30 Jahre","5–8 Mio. €","2–4 Mio. €","7–12 Mio. €"],[2600,1400,1600,1600,1826],false,"FAFBFC"),
      tRow(["Bundestagsabgeordnete","20 Jahre","2–3 Mio. €","1–2 Mio. €","3–5 Mio. €"],[2600,1400,1600,1600,1826]),
    ]
  }),
  space(),

  zitat("Die stärkste Kombination aus Geld, Einfluss und Unkontrollierbarkeit liegt nicht in der Politik — sondern im öffentlich-rechtlichen Rundfunk.","Analyse rundfunkkritik.de"),
  space(200),

  // Kapitel 6 — Vier Bremsen
  h1("VI. Was eine Bremse wirklich leisten müsste"),
  p("Die hohen Summen sind kein individuelles Versagen, sondern das Ergebnis einer Struktur, die sich selbst kontrolliert und selbst belohnt. Vier Hebel würden wirklich greifen:"),
  space(),

  h2("Hebel 1 — Transparenzpflicht"),
  p("Politiker müssen alle Bezüge offenlegen. Intendanten und Programmdirektoren nicht. Eine Transparenzpflicht nach politischem Vorbild würde das System innerhalb kurzer Zeit zur Selbstkorrektur zwingen."),
  space(60),

  h2("Hebel 2 — Gehaltsdeckelung"),
  p("Eine einfache, sofort umsetzbare Regel: Intendanten dürfen maximal das Kanzlergehalt verdienen. Programmdirektoren maximal Ministergehalt. Chefredakteure maximal Staatssekretärsgehalt. Moderatoren: maximal 500.000 Euro aus öffentlichen Mitteln."),
  space(60),

  h2("Hebel 3 — Pensionsreform"),
  p("Der größte Kostenblock ist nicht das laufende Gehalt, sondern die lebenslangen Pensionszusagen. Eine Reform nach Beamtenrecht würde die Lebenszeitvolumen der Intendanten und Direktoren annähernd halbieren."),
  space(60),

  h2("Hebel 4 — Externe Aufsicht"),
  p("Rundfunkräte werden von denselben Parteien besetzt, die der ÖRR eigentlich kontrollieren soll. Eine unabhängige Finanzkontrolle mit Veröffentlichungspflicht wäre die strukturell wirksamste Maßnahme."),
  space(200),

  // Schluss
  rule(C.dark),
  new Paragraph({ spacing:{before:200,after:80}, alignment:AlignmentType.CENTER,
    children:[r("Stand: Mai 2026  ·  rundfunkkritik.de",{size:18,italic:true,color:C.grey})] }),
  new Paragraph({ spacing:{before:0,after:0}, alignment:AlignmentType.CENTER,
    children:[r("Alle ZDF-Zahlen: geleakte Liste 2024 (Welt am Sonntag). ARD-Zahlen: Schätzungen.",
      {size:18,italic:true,color:C.grey})] }),
];


// ─── NUMBERING CONFIG ────────────────────────────────────────────────────────
const numbering = {
  config:[
    { reference:"bullets", levels:[{ level:0, format:LevelFormat.BULLET, text:"•",
      alignment:AlignmentType.LEFT,
      style:{ paragraph:{ indent:{ left:720, hanging:360 } } } }] }
  ]
};

const styles = {
  default:{ document:{ run:{ font:"Arial", size:22 } } },
  paragraphStyles:[
    { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
      run:{ size:38, bold:true, font:"Arial", color:C.dark },
      paragraph:{ spacing:{ before:480, after:160 }, outlineLevel:0,
        border:{ bottom:{ style:BorderStyle.SINGLE, size:4, color:C.mid, space:1 } } } },
    { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
      run:{ size:26, bold:true, font:"Arial", color:C.mid },
      paragraph:{ spacing:{ before:280, after:100 }, outlineLevel:1 } },
    { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
      run:{ size:22, bold:true, font:"Arial", color:C.dark },
      paragraph:{ spacing:{ before:200, after:80 }, outlineLevel:2 } },
  ]
};

const pageProps = {
  page:{
    size:{ width:11906, height:16838 },
    margin:{ top:1200, right:1200, bottom:1200, left:1200 }
  }
};

// ─── SCHREIBEN ───────────────────────────────────────────────────────────────
const doc1 = new Document({ numbering, styles,
  sections:[{ properties:pageProps, children:fallbericht }]
});
const doc2 = new Document({ numbering, styles,
  sections:[{ properties:pageProps, children:machtGeld }]
});

Promise.all([
  Packer.toBuffer(doc1),
  Packer.toBuffer(doc2)
]).then(([buf1, buf2]) => {
  fs.writeFileSync("/mnt/user-data/outputs/Fallbericht_5_Jahre_Widerstand_rundfunkkritik.docx", buf1);
  fs.writeFileSync("/mnt/user-data/outputs/Macht_Geld_ÖRR_rundfunkkritik.docx", buf2);
  console.log("OK: Beide Dokumente erstellt");
}).catch(err => {
  console.error("Fehler:", err.message);
  process.exit(1);
});
