const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  ExternalHyperlink, LevelFormat, PageBreak, PageNumber,
  Header, Footer, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ─── COLORS & BORDERS ────────────────────────────────────────────────────────
const C = { dark: "1A2F4A", mid: "2E5C8A", light: "D0E4F5", accent: "C0392B", grey: "888888", lightgrey: "F2F5F8" };
const borderNone = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const borderThin = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borderMid  = { style: BorderStyle.SINGLE, size: 4, color: C.mid };
const allBordersThin = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };
const allBordersNone = { top: borderNone, bottom: borderNone, left: borderNone, right: borderNone };

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function h(level, text, color = C.dark) {
  return new Paragraph({
    heading: level,
    spacing: { before: level === HeadingLevel.HEADING_1 ? 480 : 280, after: 120 },
    children: [new TextRun({ text, color, font: "Arial" })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 120 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text, font: "Arial", size: 22, color: opts.color || "333333", bold: opts.bold || false, italics: opts.italic || false })]
  });
}

function pRich(runs, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 120 },
    alignment: opts.align || AlignmentType.LEFT,
    children: runs
  });
}

function run(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: 22, color: opts.color || "333333", bold: opts.bold || false, italics: opts.italic || false });
}

function link(text, url) {
  return new ExternalHyperlink({
    link: url,
    children: [new TextRun({ text, font: "Arial", size: 22, color: C.mid, underline: {} })]
  });
}

function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: opts.color || "333333", bold: opts.bold || false })]
  });
}

function numbered(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: opts.color || "333333", bold: opts.bold || false })]
  });
}

function rule(color = C.mid) {
  return new Paragraph({
    spacing: { before: 240, after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color, space: 1 } },
    children: []
  });
}

function spacer(pt = 120) {
  return new Paragraph({ spacing: { before: pt, after: 0 }, children: [new TextRun("")] });
}

function infoBox(labelText, contentRuns, bgcolor = C.lightgrey) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      children: [new TableCell({
        borders: allBordersNone,
        shading: { fill: bgcolor, type: ShadingType.CLEAR },
        margins: { top: 140, bottom: 140, left: 220, right: 220 },
        width: { size: 9026, type: WidthType.DXA },
        children: [
          new Paragraph({
            spacing: { before: 0, after: 80 },
            children: [new TextRun({ text: labelText, font: "Arial", size: 20, bold: true, color: C.mid })]
          }),
          ...contentRuns.map(r => new Paragraph({
            spacing: { before: 0, after: 60 },
            children: [new TextRun({ text: r, font: "Arial", size: 20, color: "444444" })]
          }))
        ]
      })]
    })]
  });
}

function beweisBox(nummer, gruppe, beschreibung, belege) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1400, 7626],
    rows: [
      // Header row
      new TableRow({
        children: [
          new TableCell({
            borders: allBordersNone,
            shading: { fill: C.mid, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 180, right: 180 },
            width: { size: 1400, type: WidthType.DXA },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: nummer, font: "Arial", size: 28, bold: true, color: "FFFFFF" })]
            })]
          }),
          new TableCell({
            borders: allBordersNone,
            shading: { fill: C.dark, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            width: { size: 7626, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: gruppe, font: "Arial", size: 24, bold: true, color: "FFFFFF" })]
            })]
          })
        ]
      }),
      // Content row
      new TableRow({
        children: [
          new TableCell({
            borders: { top: borderNone, bottom: borderNone, left: { style: BorderStyle.SINGLE, size: 8, color: C.mid }, right: borderNone },
            shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 180, right: 180 },
            width: { size: 1400, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: "Beschreibung", font: "Arial", size: 18, bold: true, color: C.grey })]
            })]
          }),
          new TableCell({
            borders: { top: borderNone, bottom: borderNone, left: borderNone, right: borderNone },
            shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 200, right: 200 },
            width: { size: 7626, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: beschreibung, font: "Arial", size: 21, color: "333333" })]
            })]
          })
        ]
      }),
      // Belege row
      new TableRow({
        children: [
          new TableCell({
            borders: { top: borderNone, bottom: { style: BorderStyle.SINGLE, size: 2, color: C.light }, left: { style: BorderStyle.SINGLE, size: 8, color: C.mid }, right: borderNone },
            shading: { fill: C.lightgrey, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 180, right: 180 },
            width: { size: 1400, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: "Belege / Links", font: "Arial", size: 18, bold: true, color: C.grey })]
            })]
          }),
          new TableCell({
            borders: { top: borderNone, bottom: { style: BorderStyle.SINGLE, size: 2, color: C.light }, left: borderNone, right: borderNone },
            shading: { fill: C.lightgrey, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 200, right: 200 },
            width: { size: 7626, type: WidthType.DXA },
            children: belege
          })
        ]
      })
    ]
  });
}

function gerichtRow(nr, gericht, az, thema, strategie) {
  const borderR = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
  const allB = { top: borderR, bottom: borderR, left: borderR, right: borderR };
  const shade = nr % 2 === 0 ? { fill: "FAFBFC", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR };
  const cell = (text, w, bold = false) => new TableCell({
    borders: allB, shading: shade,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    width: { size: w, type: WidthType.DXA },
    children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 19, bold, color: "333333" })] })]
  });
  return new TableRow({ children: [cell(String(nr), 400, true), cell(gericht, 2200), cell(az, 1800), cell(thema, 2200), cell(strategie, 2426)] });
}

// ─── PAGE BREAK ──────────────────────────────────────────────────────────────
function pageBreak() {
  return new Paragraph({ children: [new TextRun({ break: 1 })] });
}

// ─── SCREENSHOT PLACEHOLDER ──────────────────────────────────────────────────
function screenshotPlaceholder(label, hint) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.DASHED, size: 4, color: C.grey }, bottom: { style: BorderStyle.DASHED, size: 4, color: C.grey }, left: { style: BorderStyle.DASHED, size: 4, color: C.grey }, right: { style: BorderStyle.DASHED, size: 4, color: C.grey } },
        shading: { fill: "FDF6EC", type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 220, right: 220 },
        width: { size: 9026, type: WidthType.DXA },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "📷  " + label, font: "Arial", size: 22, bold: true, color: C.accent })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: hint, font: "Arial", size: 20, italics: true, color: C.grey })] })
        ]
      })]
    })]
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT CONTENT
// ═══════════════════════════════════════════════════════════════════════════════

const children = [

  // ── TITELSEITE ──────────────────────────────────────────────────────────────
  new Paragraph({
    spacing: { before: 1200, after: 80 },
    children: [new TextRun({ text: "KLAGESCHRIFT", font: "Arial", size: 60, bold: true, color: C.dark })]
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "Rundfunkbeitrag — Modulare Beweisbausteine", font: "Arial", size: 32, color: C.mid })]
  }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "Fassung Mai 2026 · Auf Basis BVerwG 6 C 5.24 und BGH VII ZB 29/24", font: "Arial", size: 22, italics: true, color: C.grey })]
  }),
  rule(C.mid),
  infoBox("Verwendungshinweis", [
    "Dieses Dokument ist modular aufgebaut. Jeder Baustein steht für sich und kann einzeln oder",
    "kombiniert in eine Klage eingebracht werden. Gerichte reagieren zunehmend kritisch auf",
    "230-seitige Pauschalvorträge — dieser Ansatz liefert gezielt das, was das jeweilige Verfahren",
    "benötigt: Baustein A für Programmrüge, Baustein B–C für Vollstreckungsangriff,",
    "Baustein D für Datenschutz, Baustein E für den verfassungsrechtlichen Kern.",
    "",
    "Kein Rechtsanwalt — ersetzt keine anwaltliche Beratung. Alle Aktenzeichen im Dokument sind reale Verfahren."
  ]),
  spacer(200),
  pageBreak(),

  // ── INHALTSVERZEICHNIS ──────────────────────────────────────────────────────
  h(HeadingLevel.HEADING_1, "Übersicht der Bausteine"),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [900, 5626, 2500],
    rows: [
      // header
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({ borders: allBordersNone, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, width: { size: 900, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Nr.", font: "Arial", size: 20, bold: true, color: "FFFFFF" })] })] }),
          new TableCell({ borders: allBordersNone, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, width: { size: 5626, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Baustein / Thema", font: "Arial", size: 20, bold: true, color: "FFFFFF" })] })] }),
          new TableCell({ borders: allBordersNone, shading: { fill: C.dark, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, width: { size: 2500, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "Einsatzbereich", font: "Arial", size: 20, bold: true, color: "FFFFFF" })] })] }),
        ]
      }),
      ...([
        ["A", "Programmrüge — Strukturelle Defizite (BVerwG-Linie)", "Hauptklage gegen Beitragspflicht"],
        ["B", "Vollstreckungsangriff — Formelle Wirksamkeit", "Wenn Vollstreckung läuft"],
        ["C", "Zeugen & Whistleblower — Interne Kritiker", "Beweisanträge im Hauptverfahren"],
        ["D", "Datenschutz — DSGVO-Verstöße des Beitragsservice", "Separate Datenschutzklage / Zusatzrüge"],
        ["E", "Verfassungsrechtlicher Kern — Art. 3, 5, 19 GG", "Vorlage an BVerfG / Revision"],
        ["F", "Beweislast & Informationsasymmetrie", "Prozessuale Strategie in allen Verfahren"],
        ["G", "Gerichtsübersicht — 22+ Verfahren mit Aktenzeichen", "Recherchegrundlage & Verlinkung"],
        ["H", "Vollstreckungsfälle — Behörden & Muster", "Vollstreckungsstrang"],
      ]).map(([nr, thema, einsatz], i) => new TableRow({
        children: [
          new TableCell({ borders: allBordersThin, shading: { fill: i % 2 === 0 ? C.lightgrey : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 }, width: { size: 900, type: WidthType.DXA }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: nr, font: "Arial", size: 22, bold: true, color: C.mid })] })] }),
          new TableCell({ borders: allBordersThin, shading: { fill: i % 2 === 0 ? C.lightgrey : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 }, width: { size: 5626, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: thema, font: "Arial", size: 21, color: "333333" })] })] }),
          new TableCell({ borders: allBordersThin, shading: { fill: i % 2 === 0 ? C.lightgrey : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 }, width: { size: 2500, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: einsatz, font: "Arial", size: 19, italics: true, color: C.grey })] })] }),
        ]
      }))
    ]
  }),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN A — PROGRAMMRÜGE
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein A — Programmrüge: Strukturelle Defizite"),
  infoBox("Kernthese dieses Bausteins", [
    "Die Beitragspflicht ist verfassungsrechtlich an die Vorteilhaftigkeit eines vielfältigen Gesamtprogramms",
    "geknüpft. Wenn der ÖRR seinen Funktionsauftrag strukturell verfehlt, entfällt die Rechtfertigung.",
    "Grundlage: BVerwG, Urteil vom 15.10.2025 — 6 C 5.24 (Leipzig)."
  ]),
  spacer(),

  h(HeadingLevel.HEADING_2, "A-1 · Die neue BVerwG-Linie (Dreh- und Angelpunkt)"),
  pRich([run("Urteil: "), run("BVerwG, 15.10.2025 — Az. 6 C 5.24", { bold: true }), run(" · Leipzig · ", { color: C.grey }), link("→ Volltext auf bverwg.de", "https://www.bverwg.de")]),
  p("Das Bundesverwaltungsgericht hat mit diesem Urteil eine neue Messlatte gesetzt. Die alte Linie — Programmkritik spielt für die Beitragspflicht keine Rolle — gilt nicht mehr uneingeschränkt. Die neue Linie lautet:"),
  bullet("Verfassungsrechtlich besteht ein Konnex zwischen Beitragspflicht und der Vorteilhaftigkeit des Gesamtprogramms."),
  bullet("Die Beitragspflicht kann problematisch werden, wenn das Gesamtprogrammangebot über längere Zeit gröblich gegen Vielfalt und Ausgewogenheit verstößt."),
  bullet("Für einen substantiierten Vortrag braucht es in der Regel ein wissenschaftliches Gutachten — keine Einzelbeispiele."),
  bullet("Gegenstand ist das Gesamtprogramm aller ÖRR-Anstalten, nicht eine einzelne Sendung oder ein Sender."),
  spacer(),

  p("Strategischer Schlüssel: Wer einzelne Sendungen oder einzelne Themen wie Corona oder Ukraine angreift, verliert. Die Gerichte verlangen Indikatoren, Zeitraum, Korpus, Methode, Auswertung, Vergleich und Defizitkriterien über mindestens zwei Jahre.", { color: C.grey, italic: true }),
  spacer(),
  beweisBox("A-1", "BVerwG-Urteil 6 C 5.24 vom 15.10.2025",
    "Grundnorm der neuen Strategie. Verlagerung von Medienkritik zu Rechtsschutzfrage. Zitiert in fast allen Instanzverfahren 2025/2026.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ BVerwG Entscheidungssuche (Suche: 6 C 5.24)", "https://www.bverwg.de/entscheidungen/entscheidungssuche")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Beschaffung: Volltext beim Bundesverwaltungsgericht Leipzig anfordern oder über ", { italic: true, color: C.grey }), link("juris.de", "https://www.juris.de")] }),
      screenshotPlaceholder("Screenshot/PDF-Auszug einfügen: Leitsätze BVerwG 6 C 5.24", "Offizielle Veröffentlichung oder Scan der Leitsätze — Seite 1–2 des Urteils")
    ]
  ),
  spacer(200),

  h(HeadingLevel.HEADING_2, "A-2 · Das wissenschaftliche Gutachten (Vorgutachten Beitragsstopper)"),
  pRich([run("Quelle: "), run("Beitragsstopper / Freiheitskanzlei", { bold: true }), run(" · Stand Mai 2026 · "), link("→ www.beitragsstopper.de", "https://www.beitragsstopper.de")]),
  p("Das Vorgutachten der Freiheitskanzlei erfüllt sämtliche Anforderungen, die Gerichte bislang mit 'substantiierten Vortrag' gemeint haben:"),
  bullet("Gesamtprogrammanalyse — nicht einzelne Sendungen"),
  bullet("Mindestens zwei Jahre Betrachtungszeitraum"),
  bullet("Wissenschaftliche Methodik und geeignete Indikatoren"),
  bullet("Nachweis regelmäßiger und evidenter Defizite"),
  bullet("Beleg der Informationsasymmetrie (Rundfunk hat Daten, Bürger nicht)"),
  spacer(),
  beweisBox("A-2", "Vorgutachten Beitragsstopper (Freiheitskanzlei, 2026)",
    "Dieses Gutachten ist der Angriff auf die 'wissenschaftliche Hürde' der Gerichte. Es ist als Beweisantrag einreichbar und als Grundlage für sekundäre Darlegungslast nutzbar.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ www.beitragsstopper.de (Direktlink Gutachten)", "https://www.beitragsstopper.de")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Alternativ direkt bei Freiheitskanzlei anfordern: ", { italic: true, color: C.grey }), link("→ freiheitskanzlei.de", "https://www.freiheitskanzlei.de")] }),
      screenshotPlaceholder("Screenshot/PDF-Auszug einfügen: Deckblatt oder Inhaltsverzeichnis Vorgutachten", "Deckblatt + Inhaltsverzeichnis des Gutachtens als Beleg der Existenz")
    ]
  ),
  spacer(200),

  h(HeadingLevel.HEADING_2, "A-3 · Inhaltsanalysen als Beweismittel (Media Tenor / weitere)"),
  p("Gerichte verlangen keine persönliche Meinung, sondern Daten. Die folgenden Studien sind als Anlage A-3 in eine Klage einreichbar:"),
  beweisBox("A-3a", "Media Tenor: Inhaltsanalysen ARD/ZDF",
    "Wissenschaftliche Inhaltsanalysen, die über 80 % Übereinstimmung politischer Frames bei ARD/ZDF dokumentieren. Anerkannte Methodik, zitierfähig.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ mediatenor.com (Studien & Reports)", "https://www.mediatenor.com")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Buch: Roland Schatz / Media Tenor — direkt anfragen für Studiendaten", { italic: true, color: C.grey })] }),
      screenshotPlaceholder("Screenshot/Link: Media Tenor Studie oder Buchcover (C-3)", "Buchcover oder Studientitel als Belegnachweis — z.B. von mediatenor.com")
    ]
  ),
  spacer(),
  beweisBox("A-3b", "Weitere Studien: Bertelsmann, Reuters Institute, Kepplinger",
    "Unterstützende Quellen für systematische Einseitigkeit und Herdenjournalismus im deutschen ÖRR.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Reuters Institute Digital News Report (Jahresbericht, Oxford)", "https://reutersinstitute.politics.ox.ac.uk/digital-news-report")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Bertelsmann Stiftung — Publikationen Medienwandel", "https://www.bertelsmann-stiftung.de")] }),
    ]
  ),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN B — VOLLSTRECKUNGSANGRIFF
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein B — Vollstreckungsangriff: Formelle Wirksamkeit"),
  infoBox("Wann dieser Baustein greift", [
    "Sobald ein Vollstreckungsersuchen, eine Vermögensauskunft, ein Pfändungsauftrag oder",
    "eine Mahnung eingehen — sofort diesen Baustein einsetzen. Nicht abwarten.",
    "Grundlage: BGH, Beschluss vom 25.02.2026 — VII ZB 29/24."
  ]),
  spacer(),

  h(HeadingLevel.HEADING_2, "B-1 · Der BGH-Beschluss (Verantwortungszurechnung)"),
  pRich([run("Beschluss: "), run("BGH, 25.02.2026 — Az. VII ZB 29/24", { bold: true }), run(" · Bundesgerichtshof Karlsruhe · "), link("→ bundesgerichtshof.de", "https://www.bundesgerichtshof.de")]),
  p("Der BGH hat klargestellt: Auch elektronische Vollstreckungsersuchen benötigen eine konkrete persönliche Verantwortungszurechnung. Die bloße Nennung einer Intendantin oder Organbezeichnung genügt nicht. Vollautomatisierung entbindet nicht von rechtsstaatlichen Verantwortlichkeitsanforderungen."),
  spacer(),
  beweisBox("B-1", "BGH VII ZB 29/24 vom 25.02.2026",
    "Zentraler Hebel gegen maschinelle Vollstreckungsersuchen. Fragt: Wer hat konkret gehandelt? Wer hat geprüft? Wer trägt Verantwortung?",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ BGH Entscheidungssuche (Az. VII ZB 29/24)", "https://juris.bundesgerichtshof.de/cgi-bin/rechtsprechung/list.py")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Nachlesen: Markus Bönig, Substack-Essay Mai 2026 — Abschnitt IV", { italic: true, color: C.grey })] }),
      screenshotPlaceholder("Screenshot/PDF-Auszug einfügen: Leitsatz BGH VII ZB 29/24", "Beschlusstext Seite 1 oder Leitsatz der BGH-Entscheidung")
    ]
  ),
  spacer(),

  h(HeadingLevel.HEADING_2, "B-2 · AG Passau — Erste praktische Umsetzung der BGH-Linie"),
  pRich([run("Urteil: "), run("AG Passau, 08.08.2025 — Az. 4 M 5863/25", { bold: true })]),
  p("Das Amtsgericht Passau setzte die BGH-Linie erstmals praktisch um: Vollstreckungsersuchen wegen fehlender konkreter verantwortlicher Person für formunwirksam erklärt. Alle Vollstreckungsmaßnahmen einschließlich Vermögensauskunft und Eintragungsanordnung wurden aufgehoben."),
  beweisBox("B-2", "AG Passau 4 M 5863/25 vom 08.08.2025",
    "Erste operative Umsetzung der BGH-Linie. Vollstreckungsersuchen formunwirksam. Dieser Beschluss ist das Muster für alle laufenden Vollstreckungsverfahren.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Aktenzeichen: 4 M 5863/25 — Anforderung beim AG Passau oder über juristische Datenbanken", { italic: true, color: C.grey })] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Bundesamt für Justiz — Rechtsprechungsdatenbank", "https://www.buzer.de")] }),
      screenshotPlaceholder("Screenshot/PDF-Auszug einfügen: Beschluss AG Passau 4 M 5863/25", "Deckblatt oder Leitsatz des Beschlusses")
    ]
  ),
  spacer(),

  h(HeadingLevel.HEADING_2, "B-3 · Fragen, die jedes Vollstreckungsersuchen beantworten muss"),
  p("Auf Basis der BGH-Linie sind folgende Fragen an jedes Vollstreckungsersuchen zu stellen und im Schriftsatz zu formulieren:"),
  numbered("Wer hat das Vollstreckungsersuchen konkret erstellt — Name, Funktion, Datum?"),
  numbered("Wie wurde das Ersuchen elektronisch erzeugt — automatisiert oder individuell geprüft?"),
  numbered("Wer hat die inhaltliche Prüfung vorgenommen — Nachweis der Einzelfallprüfung?"),
  numbered("Wer hat elektronisch signiert — digitale Signatur, qualifiziert oder einfach?"),
  numbered("Wurde der Bescheid ordnungsgemäß bekanntgegeben — Nachweis der Zustellung?"),
  numbered("Ist der Titel bestandskräftig — oder liegt Widerspruch / Klage vor?"),
  numbered("Ist die Forderung verjährt — Prüfung § 197 BGB analog?"),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN C — ZEUGEN / WHISTLEBLOWER
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein C — Zeugen & Whistleblower: Interne Kritiker"),
  infoBox("Funktion dieses Bausteins", [
    "Aussagen ehemaliger ÖRR-Mitarbeiter sind keine Meinungen — sie sind Beweismittel.",
    "Sie belegen internen Konformitätsdruck, Regierungsnähe und strukturelle Einseitigkeit.",
    "Als Beweisantrag: 'Zum Beweis der strukturellen Einseitigkeit wird beantragt, [Name] als Zeugen zu laden.'"
  ]),
  spacer(),

  beweisBox("C-1", "Wolfgang Herles — ehemaliger ZDF-Chefredakteur",
    "Herles hat öffentlich erklärt, er und seine Kollegen seien angehalten worden, 'nicht gegen den Strom zu schwimmen'. Interner Konformitätsdruck als strukturelles Muster.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Interview Herles: 'Wir lügen nicht, aber wir verschweigen' (NachDenkSeiten 2016)", "https://www.nachdenkseiten.de/?p=31641")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Buch: Herles, 'Wir haben keine Angst' — Belege für Konformitätsdruck", { italic: true, color: C.grey })] }),
      screenshotPlaceholder("Screenshot/Link: öffentliches Interview oder Buchcover C-1", "Screenshot des NachDenkSeiten-Interviews oder Buchcover 'Wir haben keine Angst'")
    ]
  ),
  spacer(),

  beweisBox("C-2", "Ulrich Tilgner — ehemaliger ZDF-Korrespondent",
    "Tilgner berichtete über interne Zensur und Selbstzensur im Nahost-Journalismus des ZDF. Konkrete Belege für redaktionelle Eingriffe von oben.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Tilgner-Interviews (Suche: 'Ulrich Tilgner ZDF Zensur')", "https://www.google.com/search?q=Ulrich+Tilgner+ZDF+Zensur+Interview")] }),
      screenshotPlaceholder("Screenshot/Link: öffentliches Interview oder Buchcover C-2", "Screenshot eines öffentlichen Interviews von Ulrich Tilgner")
    ]
  ),
  spacer(),

  beweisBox("C-3", "Roland Schatz / Media Tenor — Wissenschaftliche Dokumentation",
    "Media Tenor dokumentiert seit Jahrzehnten mit anerkannter Methodik politische Frames, Expertenauswahl und Tonalität. Zitierfähig als wissenschaftliches Gutachten.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ mediatenor.com", "https://www.mediatenor.com")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Kontakt: Roland Schatz für Gutachten / Expertenaussage anfragen", { italic: true, color: C.grey })] }),
      screenshotPlaceholder("Screenshot/Link: öffentliches Interview oder Buchcover C-3", "Buchcover eines Media-Tenor-Reports oder Screenshot der Website mediatenor.com")
    ]
  ),
  spacer(),

  beweisBox("C-4", "Peter Hahne / Michael Meyen — Soft Ostracism",
    "Belegen das Muster der schleichenden Kaltstellung interner Kritiker (Soft Ostracism): Entzug von Ressourcen, Themen, Sendeplätzen — ohne formale Kündigung.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Meyen, Michael: 'Die Propaganda-Matrix' — Belege Soft Ostracism", "https://www.google.com/search?q=Michael+Meyen+Propaganda+Matrix")] }),
      screenshotPlaceholder("Screenshot/Link: öffentliches Interview oder Buchcover C-4", "Buchcover oder Screenshot eines öffentlichen Interviews Hahne / Meyen")
    ]
  ),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN D — DATENSCHUTZ
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein D — Datenschutz: DSGVO-Verstöße"),
  infoBox("Besonderheit: Separater Klageweg", [
    "Datenschutzverstöße können unabhängig von der Beitragsfrage geltend gemacht werden.",
    "Zuständig: Bayerisches Landesamt für Datenschutzaufsicht (BayLDA) oder direkt Verwaltungsgericht.",
    "Dieser Baustein ergänzt jeden anderen Baustein und ist separat einreichbar."
  ]),
  spacer(),

  beweisBox("D-1", "Meldedatenabgleich — Rechtsgrundlage und Transparenzpflicht",
    "Der Beitragsservice führt automatisierte Meldedatenabgleiche durch. Rechtsgrundlage: § 14 RBStV. Fraglich: Erfüllt er die DSGVO-Anforderungen an Zweckbindung, Transparenz, Verhältnismäßigkeit?",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ § 14 RBStV (Gesetzeswortlaut auf gesetze-im-internet.de)", "https://www.gesetze-im-internet.de")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("Rüge: Fehlende DSGVO-Datenschutz-Folgenabschätzung (Art. 35 DSGVO)", { italic: true, color: C.grey })] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Beschwerde beim BayLDA: www.lda.bayern.de", "https://www.lda.bayern.de")] }),
    ]
  ),
  spacer(),

  beweisBox("D-2", "Automatisierte Bescheide — Verbot nach Art. 22 DSGVO",
    "Vollautomatisierte Entscheidungen ohne menschliche Prüfung verstoßen gegen Art. 22 DSGVO, sofern sie den Betroffenen erheblich beeinflussen. Rundfunkbeitragsbescheide sind erhebliche Eingriffe.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ Art. 22 DSGVO (Datenschutz-Grundverordnung)", "https://dsgvo-gesetz.de/art-22-dsgvo/")] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [run("VG Würzburg W 3 S 25.2191 trennt explizit: Beitragsbescheid / Datenverarbeitung / Unterlassung", { italic: true, color: C.grey })] }),
    ]
  ),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN E — VERFASSUNGSRECHTLICHER KERN
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein E — Verfassungsrechtlicher Kern"),
  infoBox("Für Musterverfahren und Revisionen", [
    "Dieser Baustein gehört in Musterverfahren, nicht in Massenverfahren.",
    "Massenverfahren: Baustein A/B, kleiner Streitwert, nur Bescheidaufhebung.",
    "Musterverfahren: Baustein E mit Vorlageantrag an BVerfG — BVerfG-Termin 23. Juni 2026 ist aktuell!"
  ]),
  spacer(),

  h(HeadingLevel.HEADING_2, "E-1 · Asymmetrische Aufsicht (Art. 3 GG)"),
  p("Freie Medien unterliegen der Kontrolle durch Landesmedienanstalten und können sanktioniert werden. Der ÖRR unterliegt nur internen, politisch besetzten Rundfunkräten. Diese strukturelle Ungleichbehandlung verletzt Art. 3 GG — gleiche Anforderungen an Vielfalt, ungleiche Kontrolle."),
  spacer(),

  h(HeadingLevel.HEADING_2, "E-2 · Rechtsschutzgarantie (Art. 19 Abs. 4 GG)"),
  p("Art. 19 Abs. 4 GG garantiert effektiven Rechtsschutz. Wenn Gerichte wissenschaftliche Gesamtanalysen verlangen und zugleich die Datenoffenlegung verweigern, läuft dieser Rechtsschutz leer. Der Bürger soll ein System beweisen, zu dessen Daten er keinen Zugang hat."),
  p("Argumentation: Sekundäre Darlegungslast — wenn der Rundfunk die Daten hat und verweigert, muss die Darlegungslast prozessual angepasst werden. Dann muss das Gericht nach § 86 VwGO aufklären. Andernfalls: Art. 19 Abs. 4 GG ausgehöhlt.", { color: C.grey, italic: true }),
  spacer(),

  h(HeadingLevel.HEADING_2, "E-3 · Rundfunkfreiheit vs. Staatsferne (Art. 5 GG)"),
  p("Art. 5 GG garantiert nicht nur Rundfunkfreiheit, sondern schützt auch vor staatlichem Einfluss auf den Rundfunk. Politisch besetzte Rundfunkräte, Intendantenernennungen durch Landesregierungen und Näheverhältnisse zur Politik verletzen das Staatsferneprinzip des BVerfG."),
  beweisBox("E-3", "BVerfG 1 BvF 1/11 — ZDF-Staatsvertrag (2014)",
    "Das Bundesverfassungsgericht hat 2014 klar formuliert: Der ÖRR muss staatsfern sein. Politisch besetzte Gremien verstoßen gegen dieses Gebot. Dieses Urteil ist Fundament für alle Staatsferneargumente.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ BVerfG 1 BvF 1/11 vom 25.03.2014 (Urteil ZDF-Staatsvertrag)", "https://www.bundesverfassungsgericht.de/SharedDocs/Entscheidungen/DE/2014/03/fs20140325_1bvf000111.html")] }),
      screenshotPlaceholder("Screenshot/PDF-Auszug einfügen: Leitsätze BVerfG 1 BvF 1/11", "Leitsätze des ZDF-Staatsvertragsurteils von 2014 — Seite 1–2")
    ]
  ),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN F — BEWEISLAST & INFORMATIONSASYMMETRIE
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein F — Beweislast & Informationsasymmetrie"),
  p("Dies ist der prozessuale Schlüsselbaustein. Er passt in jedes Verfahren."),
  spacer(),

  h(HeadingLevel.HEADING_2, "F-1 · Sekundäre Darlegungslast"),
  p("Der BGH-Baustein vom 09.03.2026, Az. VI ZR 335/24, hat diese Linie gestärkt: Wenn nur eine Partei Zugang zu entscheidenden Informationen hat, darf die andere nicht an Beweisnot scheitern. Übertragen auf Rundfunkverfahren:"),
  bullet("Der Bürger sieht nur das ausgestrahlte Endprodukt."),
  bullet("Der Rundfunk kontrolliert Archive, Daten, Auswahlmechanismen, interne Qualitätskontrolle, Redaktionsleitlinien."),
  bullet("Verweigert der Rundfunk Datenzugang: strukturelle Beweisnot."),
  bullet("Dann muss das Gericht nach § 86 VwGO aufklären — oder die Darlegungslast verschiebt sich."),
  spacer(),
  beweisBox("F-1", "BGH VI ZR 335/24 vom 09.03.2026",
    "Sekundäre Darlegungslast bei Informationsasymmetrie. Direkt auf Rundfunkverfahren übertragbar: Der Rundfunk hat die Daten, der Bürger nicht.",
    [
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [link("→ BGH Rechtsprechung (Az. VI ZR 335/24)", "https://juris.bundesgerichtshof.de/cgi-bin/rechtsprechung/list.py")] }),
    ]
  ),
  spacer(),

  h(HeadingLevel.HEADING_2, "F-2 · Prozessuale Doppelstrategie"),
  p("Aus dem Bönig-Essay (Mai 2026) — die Kernstrategie für alle Verfahren:"),
  bullet("Massenverfahren klein halten: Nur Bescheidaufhebung, niedriger Streitwert, keine überladenen Systemanträge."),
  bullet("Musterverfahren groß führen: Gutachten, Datenzugang, Rohdaten, Beweisanträge, sekundäre Darlegungslast, Art. 19 Abs. 4 GG, § 86 VwGO."),
  bullet("Nie beides mischen — das ist der häufigste Fehler der bisherigen 230-Seiten-Klagen."),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN G — GERICHTSÜBERSICHT
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein G — Gerichtsübersicht: 22+ Verfahren"),
  p("Alle genannten Aktenzeichen stammen aus dem Bönig-Essay (Substack, Mai 2026). Stand: Mai 2026."),
  spacer(),

  h(HeadingLevel.HEADING_2, "G-1 · Programmrüge-Verfahren (Instanzgerichte)"),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [400, 2200, 1800, 2200, 2426],
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["#", "Gericht", "Aktenzeichen", "Thema", "Strateg. Kommentar"].map((t, i) => new TableCell({
          borders: allBordersNone,
          shading: { fill: C.dark, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          width: { size: [400,2200,1800,2200,2426][i], type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 19, bold: true, color: "FFFFFF" })] })]
        }))
      }),
      gerichtRow(1, "BVerwG Leipzig", "6 C 5.24", "Grundnorm — Programmrüge", "Fundament aller Verfahren 2025/26"),
      gerichtRow(2, "BayVGH München", "7 BV 25.2291", "BR-Abwehr / Darlegungshürde", "BR nutzt BVerwG als Schutzwall"),
      gerichtRow(3, "VG Augsburg", "Hinweis 26.01.2026", "Gutachten-Schwelle", "Operationalisiert BVerwG-Formel"),
      gerichtRow(4, "VG Regensburg", "RN 3 K 25.291", "Eilverfahren Vollstreckung", "Rechtsschutzlücke bei laufender Vollstreckung"),
      gerichtRow(5, "VG Stuttgart", "14 K 1203/25", "Darlegungslast / Klägeraufgabe", "SWR-Parallelentscheidung"),
      gerichtRow(6, "VG Stuttgart", "14 K 15357/25", "VGH-BW Druckmittel", "Psychologischer Rückzugsdruck"),
      gerichtRow(7, "VG Düsseldorf", "27 K 9336/24", "Rücknahmeempfehlung § 84", "Vorzeigen des Ausgangs vorab"),
      gerichtRow(8, "VG Weimar", "3 K 20/26 We", "Keine sek. Darlegungslast", "Dogmatisch sauber — praktisch brutal"),
      gerichtRow(9, "VG Potsdam", "VG 11 K 3620/24", "Gesamtprogramm / Kompensation", "Fragt nach mündl. Verhandlung"),
      gerichtRow(10, "VG Saarland", "1 K 1904/25", "Gerichtsbescheid angekündigt", "Frühe Post-BVerwG-Umsetzung"),
      gerichtRow(11, "VG Freiburg", "1 K 5977/24", "Gesamtprogramm-Formel", "Standardformulierung BVerwG-Linie"),
      gerichtRow(12, "VG Hannover", "7 A 4178/24", "NDR widerspricht Ruhen", "NDR verhindert Bündelung"),
      gerichtRow(13, "VG Berlin", "VG 27 K 343/24", "Informationsasymmetrie RBB", "RBB: öffentlich zugänglich = kein Problem"),
      gerichtRow(14, "VG Berlin", "VG 27 K 460/24", "Datenzugang / Asymmetrie", "Gegenargument wird eingeführt"),
      gerichtRow(15, "VG Würzburg", "W 3 K 25.2191", "Trennung der Streitgegenstände", "Gerichte zerlegen Komplexklagen"),
      gerichtRow(16, "VG Würzburg", "W 3 S 25.2191", "Unterlassung Datenverarbeitung", "DSGVO-Strang separat"),
      gerichtRow(17, "VG Oldenburg", "15 A 5793/25", "Amtsermittlung ausgeschlossen", "BVerwG-Anforderungen übernommen"),
      gerichtRow(18, "VG Lüneburg", "3 A 113/20", "Kompensations-Gesamtprogramm", "Frühe NDR-Entscheidung"),
      gerichtRow(19, "VG Trier", "(Az. nicht lesbar)", "Frühe Standardisierung", "Post-BVerwG schnelle Linie"),
      gerichtRow(20, "VG Wiesbaden", "2 K 202/26.WI", "HR — klassische Abwehr", "Alt + neu kombiniert"),
      gerichtRow(21, "VG Kassel", "(Az. nicht lesbar)", "Widerspruch offen beschrieben", "Art.-19-Abs.-4-Hebel"),
      gerichtRow(22, "VG Meiningen", "(Az. nicht lesbar)", "MDR-Linie: keine Daten", "Beweisnot-These gestärkt"),
    ]
  }),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // BAUSTEIN H — VOLLSTRECKUNGSFÄLLE
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Baustein H — Vollstreckungsfälle: Behörden & Muster"),
  p("Diese Fälle zeigen das gemeinsame Muster: Behörden handeln, wollen aber nicht verantwortlich sein."),
  spacer(),

  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [400, 2400, 1800, 4426],
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["#", "Behörde / Gericht", "Aktenzeichen", "Kernargument der Behörde"].map((t, i) => new TableCell({
          borders: allBordersNone,
          shading: { fill: C.accent, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          width: { size: [400,2400,1800,4426][i], type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 19, bold: true, color: "FFFFFF" })] })]
        }))
      }),
      ...[
        [1, "BGH (Vollstreckung)", "VII ZB 29/24 (25.02.2026)", "Elektronische Ersuchen brauchen persönliche Verantwortung — Organfiktion reicht nicht"],
        [2, "AG Passau", "4 M 5863/25 (08.08.2025)", "Vollstreckung formunwirksam — keine konkrete verantwortliche Person benannt"],
        [3, "AG Augsburg", "02 M 4262/26", "Frage: Ist Verfahren abgeschlossen? — maschinell gezeichneter Auftrag problematisch"],
        [4, "LG Kempten", "42 T 408/25 (17.03.2026)", "Namens-/Organangabe reicht nicht — Frage echter Verantwortungsübernahme"],
        [5, "GV Viersen / NRW", "Az. 176853416", "BGH betreffe nur Bayern — NRW vollstreckt nach VwVG NRW"],
        [6, "Stadt Koblenz", "21/VST-NO 204534 2026/6393", "BGH betreffe nur SWR/Bayern — Zuständigkeitsflucht"],
        [7, "Gemeinde Neuenhagen", "BGM_AA_R_0012-26", "BGH sei Verfahrensfehler — Vollstreckung schon 2023 abgeschlossen"],
        [8, "Saale-Holzland-Kreis", "(Thüringen)", "Amtshilfe MDR — elektronische Schnittstelle — sachbearbeiterlos"],
        [9, "Bergisch Gladbach", "DR II 265/26", "Vollstreckungsersuchen / Verantwortlichkeit / BGH-Linie als Einwand"],
        [10, "Neuss/Dormagen WDR", "DR II 333/26", "Maschinenraum sichtbar: wer ist Intendantin — echte Übernahme oder Symbolik?"],
      ].map(([nr, beh, az, arg], i) => new TableRow({
        children: [
          new TableCell({ borders: allBordersThin, shading: { fill: i%2===0?C.lightgrey:"FFFFFF", type: ShadingType.CLEAR }, margins:{top:80,bottom:80,left:120,right:120}, width:{size:400,type:WidthType.DXA}, children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:String(nr),font:"Arial",size:19,bold:true,color:C.accent})]})] }),
          new TableCell({ borders: allBordersThin, shading: { fill: i%2===0?C.lightgrey:"FFFFFF", type: ShadingType.CLEAR }, margins:{top:80,bottom:80,left:120,right:120}, width:{size:2400,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:beh,font:"Arial",size:19,bold:false,color:"333333"})]})] }),
          new TableCell({ borders: allBordersThin, shading: { fill: i%2===0?C.lightgrey:"FFFFFF", type: ShadingType.CLEAR }, margins:{top:80,bottom:80,left:120,right:120}, width:{size:1800,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:az,font:"Arial",size:17,color:C.grey})]})]}),
          new TableCell({ borders: allBordersThin, shading: { fill: i%2===0?C.lightgrey:"FFFFFF", type: ShadingType.CLEAR }, margins:{top:80,bottom:80,left:120,right:120}, width:{size:4426,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:arg,font:"Arial",size:19,color:"333333"})]})] }),
        ]
      }))
    ]
  }),
  spacer(200),
  pageBreak(),

  // ═══════════════════════════════════════════════════════════════════════════
  // ANHANG — QUELLEN & NÄCHSTE SCHRITTE
  // ═══════════════════════════════════════════════════════════════════════════
  h(HeadingLevel.HEADING_1, "Anhang — Quellen, Links & nächste Schritte"),

  h(HeadingLevel.HEADING_2, "Primärquellen (Gerichtsurteile)"),
  pRich([link("BVerwG 6 C 5.24 (15.10.2025) → bverwg.de", "https://www.bverwg.de/entscheidungen/entscheidungssuche")]),
  pRich([link("BGH VII ZB 29/24 (25.02.2026) → bundesgerichtshof.de", "https://juris.bundesgerichtshof.de/cgi-bin/rechtsprechung/list.py")]),
  pRich([link("BVerfG 1 BvF 1/11 (ZDF-Staatsvertrag 2014) → bundesverfassungsgericht.de", "https://www.bundesverfassungsgericht.de/SharedDocs/Entscheidungen/DE/2014/03/fs20140325_1bvf000111.html")]),
  spacer(),

  h(HeadingLevel.HEADING_2, "Sekundärquellen"),
  pRich([link("Bönig, Markus: Essay 'Der Rundfunkbeitrag vor Gericht' (Substack, Mai 2026)", "https://markusboenigsubstack.substack.com")]),
  pRich([link("Beitragsstopper / Freiheitskanzlei → www.beitragsstopper.de", "https://www.beitragsstopper.de")]),
  pRich([link("Media Tenor (Schatz) → www.mediatenor.com", "https://www.mediatenor.com")]),
  spacer(),

  h(HeadingLevel.HEADING_2, "Nächste Schritte — Pendeliste"),
  numbered("Prof. Rieck und Roland Schatz anschreiben (diese Woche)"),
  numbered("Banner 6 auf Social Media posten — BVerfG-Termin 23. Juni 2026"),
  numbered("Artikel zu BVerwG + BVerfG-Ausblick schreiben (zeitkritisch — vor 23. Juni!)"),
  numbered("Postfach-Adresse ins Impressum eintragen"),
  numbered("Downloads-Bereich: Klageschrift-PDF + DOCX ergänzen"),
  numbered("Sitemap bei Google Search Console einreichen"),
  numbered("Favicon für Tab-Leiste einbauen"),
  numbered("Link Bräutigam (ÖRR-Kritik) — toter Link ersetzen oder streichen"),
  spacer(),
  rule(C.accent),
  p("Stand: Mai 2026 · rundfunkkritik.de · Kein Rechtsanwalt. Kein Ersatz für anwaltliche Beratung.", { color: C.grey, italic: true, align: AlignmentType.CENTER }),
];

// ─── DOCUMENT ────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 40, bold: true, font: "Arial", color: C.dark },
        paragraph: { spacing: { before: 480, after: 160 }, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.mid, space: 1 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: C.mid },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.mid, space: 1 } },
          spacing: { before: 0, after: 160 },
          children: [
            new TextRun({ text: "Klageschrift Rundfunkbeitrag — Modulare Bausteine  |  Stand Mai 2026", font: "Arial", size: 18, color: C.grey }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.mid, space: 1 } },
          spacing: { before: 160, after: 0 },
          children: [
            new TextRun({ text: "rundfunkkritik.de  |  Kein Rechtsanwalt — kein Ersatz für anwaltliche Beratung  |  Seite ", font: "Arial", size: 18, color: C.grey }),
            new TextRun({ children: [{ type: "page" }], font: "Arial", size: 18, color: C.grey }),
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/Klageschrift_Rundfunk_Bausteine_Mai2026.docx", buffer);
  console.log("OK: Dokument erstellt");
}).catch(err => {
  console.error("Fehler:", err.message);
  process.exit(1);
});
