"use strict";
const express = require("express");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber,
  TabStopType, TabStopPosition, PageOrientation, SectionType
} = require("docx");

const app  = express();
app.use(express.json({ limit: "10mb" }));

// ── Embedded logo (base64 PNG, navy background) ──────────────────────────
const LOGO_B64 = "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";
const logoData = Buffer.from(LOGO_B64, "base64");

// ── Palette ──────────────────────────────────────────────────────────────
const C = {
  orange:     "FF6734", navy:      "0F1F38", ink:       "1C1C2E",
  mid:        "4A4A6A", muted:     "7A7A96", rule:      "E2E2ED",
  surfaceAlt: "F4F4F0", amberText: "8A5A00", amberBg:   "FDF5E0",
  greenText:  "1A6B4A", greenBg:   "EAF5EF", white:     "FFFFFF",
};

// ── Page geometry ─────────────────────────────────────────────────────────
// A4 Portrait  content width = 11906 - 1440*2 = 9026 DXA
// A4 Landscape content width = 16838 - 1440*2 = 13958 DXA
const PG  = { W: 9026  };
const PGL = { W: 13958 };

// ── Border helpers ────────────────────────────────────────────────────────
const noBorder  = { style: BorderStyle.NONE, size: 0, color: "auto" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
function solidBorder(color, size = 4) {
  return { style: BorderStyle.SINGLE, size, color };
}

// ── Shading ───────────────────────────────────────────────────────────────
function shade(fill) { return { fill, type: ShadingType.CLEAR, color: "auto" }; }

// ── Margins ───────────────────────────────────────────────────────────────
const CM  = { top: 80,  bottom: 80,  left: 120, right: 120 };
const CMW = { top: 120, bottom: 120, left: 160, right: 160 };

// ── Text helpers ──────────────────────────────────────────────────────────
function run(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: 20, color: C.ink, ...opts });
}
function para(children, opts = {}) {
  if (typeof children === "string") children = [run(children)];
  return new Paragraph({ children, spacing: { after: 80 }, ...opts });
}
function emptyPara() { return new Paragraph({ children: [run("")], spacing: { after: 0 } }); }

// ── Section heading ───────────────────────────────────────────────────────
function sectionHeading(text) {
  return [new Paragraph({
    children: [new TextRun({ text, font: "Georgia", size: 32, bold: true, color: C.navy })],
    spacing: { before: 480, after: 160 },
    border: { bottom: solidBorder(C.rule, 4) },
  })];
}

// ── Sub-heading ───────────────────────────────────────────────────────────
function subHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), font: "Arial", size: 16,
      bold: true, color: C.navy, characterSpacing: 40 })],
    spacing: { before: 200, after: 80 },
  });
}

// ── Claim block (left orange border) ─────────────────────────────────────
function claimBlock(label, text) {
  return new Table({
    width: { size: PG.W, type: WidthType.DXA },
    columnWidths: [60, PG.W - 60],
    borders: noBorders,
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: { top: noBorder, bottom: noBorder, left: solidBorder(C.orange, 12), right: noBorder },
          shading: shade(C.white), margins: { top: 0, bottom: 0, left: 0, right: 0 },
          width: { size: 60, type: WidthType.DXA },
          children: [emptyPara()],
        }),
        new TableCell({
          borders: noBorders, shading: shade(C.white), margins: CMW,
          width: { size: PG.W - 60, type: WidthType.DXA },
          children: [
            new Paragraph({
              children: [new TextRun({ text: label.toUpperCase(),
                font: "Arial", size: 16, bold: true, color: C.orange, characterSpacing: 40 })],
              spacing: { after: 80 },
            }),
            new Paragraph({
              children: [new TextRun({ text, font: "Arial", size: 19, color: C.ink, italics: true })],
              spacing: { after: 0 },
            }),
          ],
        }),
      ],
    })],
  });
}

// ── Summary card table ────────────────────────────────────────────────────
function summaryCardTable(cards) {
  const colW = Math.floor(PG.W / cards.length);
  return new Table({
    width: { size: PG.W, type: WidthType.DXA },
    columnWidths: cards.map(() => colW),
    borders: noBorders,
    rows: [new TableRow({
      children: cards.map(c => new TableCell({
        borders: {
          top: solidBorder(C.orange, 8), bottom: solidBorder(C.rule, 4),
          left: noBorder, right: solidBorder(C.rule, 4),
        },
        shading: shade(c.highlight ? C.amberBg : C.white),
        margins: CMW,
        width: { size: colW, type: WidthType.DXA },
        children: [
          new Paragraph({
            children: [new TextRun({ text: c.label.toUpperCase(),
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
            spacing: { after: 60 },
          }),
          new Paragraph({
            children: [new TextRun({ text: c.value, font: "Arial",
              size: c.small ? 20 : 28, bold: true,
              color: c.highlight ? C.amberText : C.navy })],
            spacing: { after: 0 },
          }),
        ],
      })),
    })],
  });
}

// ── Mapping item ──────────────────────────────────────────────────────────
function mappingItem(num, featureText, conclusion, rationale) {
  return new Table({
    width: { size: PG.W, type: WidthType.DXA },
    columnWidths: [400, PG.W - 400],
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: solidBorder(C.rule, 4), bottom: noBorder,
              left: solidBorder(C.rule, 4), right: noBorder },
            shading: shade(C.white), margins: CMW,
            width: { size: 400, type: WidthType.DXA },
            verticalAlign: VerticalAlign.TOP,
            children: [new Paragraph({
              children: [new TextRun({ text: String(num), font: "Arial",
                size: 24, bold: true, color: C.white })],
              alignment: AlignmentType.CENTER, shading: shade(C.navy), spacing: { after: 0 },
            })],
          }),
          new TableCell({
            borders: { top: solidBorder(C.rule, 4), bottom: noBorder,
              left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4) },
            shading: shade(C.white), margins: CMW,
            width: { size: PG.W - 400, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: featureText, font: "Arial",
                size: 20, bold: true, color: C.ink })],
              spacing: { after: 0 },
            })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: solidBorder(C.rule, 4), right: noBorder },
            shading: shade(C.surfaceAlt), margins: CMW,
            width: { size: 400, type: WidthType.DXA },
            children: [emptyPara()],
          }),
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4) },
            shading: shade(C.surfaceAlt), margins: CMW,
            width: { size: PG.W - 400, type: WidthType.DXA },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "Conclusion:  ", font: "Arial", size: 19, bold: true, color: C.ink }),
                  new TextRun({ text: conclusion, font: "Arial", size: 19, color: C.mid }),
                ],
                spacing: { after: 80 },
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: "Brief Rationale:  ", font: "Arial", size: 19, bold: true, color: C.ink }),
                  new TextRun({ text: rationale, font: "Arial", size: 19, color: C.mid }),
                ],
                spacing: { after: 0 },
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// ── Justification panel ───────────────────────────────────────────────────
function justificationPanel(text, W = PG.W) {
  return new Table({
    width: { size: W - 80, type: WidthType.DXA },
    columnWidths: [48, W - 128],
    borders: noBorders,
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: { top: noBorder, bottom: noBorder, left: solidBorder(C.orange, 10), right: noBorder },
          shading: shade(C.white), width: { size: 48, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [emptyPara()],
        }),
        new TableCell({
          borders: noBorders, shading: shade(C.white),
          width: { size: W - 128, type: WidthType.DXA }, margins: CM,
          children: [
            new Paragraph({
              children: [new TextRun({ text: "ESSENTIALITY JUSTIFICATION",
                font: "Arial", size: 15, bold: true, color: C.orange, characterSpacing: 40 })],
              spacing: { after: 80 },
            }),
            new Paragraph({
              children: [new TextRun({ text, font: "Arial", size: 19, color: C.mid })],
              spacing: { after: 0 },
            }),
          ],
        }),
      ],
    })],
  });
}

// ── Analysis paragraphs ───────────────────────────────────────────────────
function analysisParagraphs(interpretation, mappingDetail, differences, opinion) {
  // mappingDetail may be a string (newline-separated) or array
  const lines = Array.isArray(mappingDetail)
    ? mappingDetail
    : String(mappingDetail || "").split(/\n\n+/).filter(Boolean);

  return [
    subHeading("Interpretation"),
    new Paragraph({ children: [run(interpretation, { size: 19, color: C.mid })], spacing: { after: 120 } }),
    subHeading("Mapping Summary"),
    ...lines.map(line => new Paragraph({
      children: [run(String(line).replace(/\*\*/g, ""), { size: 19, color: C.mid })],
      spacing: { after: 80 },
    })),
    subHeading("Differences"),
    new Paragraph({ children: [run(differences, { size: 19, color: C.mid })], spacing: { after: 120 } }),
    subHeading("Overall Opinion"),
    new Paragraph({ children: [run(opinion, { size: 19, color: C.mid })], spacing: { after: 160 } }),
  ];
}

// ── Excerpt item ──────────────────────────────────────────────────────────
function excerptItem(num, ref, heading, bodyLines, W = PG.W) {
  const labelW = 1100;
  const refW   = W - labelW;
  const bodyChildren = [];
  if (heading) {
    bodyChildren.push(new Paragraph({
      children: [new TextRun({ text: heading, font: "Arial", size: 16,
        bold: true, color: C.navy, characterSpacing: 30 })],
      spacing: { after: 80 },
    }));
  }
  bodyLines.forEach(line => {
    bodyChildren.push(new Paragraph({
      children: [new TextRun({ text: String(line), font: "Courier New", size: 15, color: C.mid })],
      spacing: { after: 60 },
    }));
  });
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [labelW, refW],
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
              left: solidBorder(C.rule,4), right: noBorder },
            shading: shade(C.surfaceAlt), margins: CM, width: { size: labelW, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: "Excerpt " + num, font: "Arial",
                size: 17, bold: true, color: C.navy, characterSpacing: 30 })],
              spacing: { after: 0 },
            })],
          }),
          new TableCell({
            borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
              left: noBorder, right: solidBorder(C.rule,4) },
            shading: shade(C.surfaceAlt), margins: CM, width: { size: refW, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: ref, font: "Courier New", size: 15, color: C.muted })],
              alignment: AlignmentType.RIGHT, spacing: { after: 0 },
            })],
          }),
        ],
      }),
      new TableRow({
        children: [new TableCell({
          columnSpan: 2,
          borders: { top: noBorder, bottom: solidBorder(C.rule,4),
            left: solidBorder(C.rule,4), right: solidBorder(C.rule,4) },
          shading: shade("FAFAFA"), margins: CM, width: { size: W, type: WidthType.DXA },
          children: bodyChildren,
        })],
      }),
    ],
  });
}

// ── Feature block (landscape claim chart) ────────────────────────────────
function featureBlock(num, featureText, disclosure, essentiality,
                      analysisChildren, excerptTables, W = PGL.W) {
  const verdictText = disclosure + "  ·  " + essentiality;
  const colW = Math.floor(W / 2);

  const leftChildren = [
    new Paragraph({
      children: [new TextRun({ text: "ANALYSIS", font: "Arial", size: 17,
        bold: true, color: C.muted, characterSpacing: 60 })],
      border: { bottom: solidBorder(C.rule, 4) },
      spacing: { after: 140, before: 0 },
    }),
    ...analysisChildren,
  ];

  const rightChildren = [
    new Paragraph({
      children: [new TextRun({ text: "CITED STANDARD EXCERPTS", font: "Arial",
        size: 17, bold: true, color: C.muted, characterSpacing: 60 })],
      border: { bottom: solidBorder(C.rule, 4) },
      spacing: { after: 140, before: 0 },
    }),
    ...excerptTables.flatMap(t => [t, emptyPara()]),
  ];

  return [
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [400, W - 400],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: noBorders, shading: shade(C.navy),
            margins: { top: 120, bottom: 120, left: 160, right: 80 },
            width: { size: 400, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              children: [new TextRun({ text: String(num), font: "Arial",
                size: 28, bold: true, color: C.white })],
              alignment: AlignmentType.CENTER, spacing: { after: 0 },
            })],
          }),
          new TableCell({
            borders: noBorders, shading: shade(C.navy),
            margins: { top: 120, bottom: 120, left: 80, right: 160 },
            width: { size: W - 400, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              children: [new TextRun({ text: featureText, font: "Arial",
                size: 19, color: "DDDDDD", italics: true })],
              spacing: { after: 0 },
            })],
          }),
        ],
      })],
    }),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [W],
      borders: noBorders,
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: noBorder, bottom: solidBorder(C.rule, 4), left: noBorder, right: noBorder },
          shading: shade(C.amberBg),
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          width: { size: W, type: WidthType.DXA },
          children: [new Paragraph({
            children: [new TextRun({ text: verdictText, font: "Arial",
              size: 17, bold: true, color: C.amberText, characterSpacing: 30 })],
            spacing: { after: 0 },
          })],
        })],
      })],
    }),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [colW, W - colW],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4) },
            shading: shade(C.white), margins: CMW,
            width: { size: colW, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
            children: leftChildren,
          }),
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: noBorder, right: solidBorder(C.rule, 4) },
            shading: shade(C.white), margins: CMW,
            width: { size: W - colW, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
            children: rightChildren,
          }),
        ],
      })],
    }),
    emptyPara(),
  ];
}

// ── Header / Footer ───────────────────────────────────────────────────────
function makeHeader(contentW) {
  return new Header({
    children: [
      new Table({
        width: { size: contentW, type: WidthType.DXA },
        columnWidths: [2800, contentW - 2800],
        borders: noBorders,
        rows: [new TableRow({
          children: [
            new TableCell({
              borders: noBorders, shading: shade(C.navy),
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: 2800, type: WidthType.DXA },
              children: [new Paragraph({
                children: [new ImageRun({
                  type: "png", data: logoData,
                  transformation: { width: 160, height: 52 },
                  altText: { title: "IPMIND", description: "IPMIND logo", name: "logo" },
                })],
                spacing: { after: 0 },
              })],
            }),
            new TableCell({
              borders: noBorders, shading: shade(C.navy),
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: contentW - 2800, type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({
                children: [new TextRun({ text: "CONFIDENTIAL", font: "Arial",
                  size: 16, color: "888888", characterSpacing: 80 })],
                alignment: AlignmentType.RIGHT, spacing: { after: 0 },
              })],
            }),
          ],
        })],
      }),
      new Paragraph({
        children: [run("")],
        border: { bottom: solidBorder(C.orange, 12) },
        spacing: { after: 0, before: 0 },
      }),
    ],
  });
}

function makeFooter() {
  return new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: "ipmind.ai", font: "Arial", size: 16, color: C.muted }),
        new TextRun({ text: "\t", font: "Arial", size: 16 }),
        new TextRun({ text: "Page ", font: "Arial", size: 16, color: C.muted }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: C.muted }),
        new TextRun({ text: " of ", font: "Arial", size: 16, color: C.muted }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: C.muted }),
      ],
      border: { top: solidBorder(C.rule, 4) },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      spacing: { before: 120, after: 0 },
    })],
  });
}

// ── Disclaimer ────────────────────────────────────────────────────────────
function disclaimerSection() {
  const items = [
    ["Preliminary and Informational Nature:", "The present work product was generated using a prototype AI model and is provided for informational purposes only. It does not constitute a legal or technical opinion regarding the essentiality or non-essentiality of any patent claim to any technical standard."],
    ["Scope of Analysis:", "The analysis is limited to the individual patent claim(s) identified in the chart and does not take into account the full patent specification, including the description and drawings."],
    ["Referencing of Standards:", "Where citations to section numbers, table numbers, or figure numbers in a technical standard are provided, they are included for convenience only and should not be relied upon as authoritative without verification against the official version of the standard."],
    ["Interpretation of Standards:", "References to technical standards are based on publicly available documents. Figures and diagrams from such standards are not reproduced; instead, any associated visual content is paraphrased using descriptive language."],
    ["Subjectivity of Essentiality:", "Determinations of potential alignment between a patent claim and a standard may depend on how specific terms or functional steps are construed. This assessment is inherently interpretive and does not reflect a consensus view or judicial determination."],
    ["Implementation Considerations:", "The presence of a feature in a standard does not imply that all compliant implementations necessarily use that feature."],
    ["Alternative Solutions:", "Standards may include multiple options or alternative techniques to achieve similar functionality. A given patent claim may correspond to one such option, but not to others that are also compliant with the standard."],
    ["Legal Proceedings:", "In the context of litigation, essentiality determinations typically require expert testimony, claim construction under applicable law, and examination of implementation evidence. The present assessment should not be relied upon for litigation or licensing negotiation without further professional review."],
  ];
  return [
    ...sectionHeading("Disclaimer"),
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [PG.W],
      borders: noBorders,
      rows: [new TableRow({
        children: [new TableCell({
          borders: {
            top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
            left: solidBorder(C.rule,4), right: solidBorder(C.rule,4),
          },
          shading: shade(C.surfaceAlt), margins: CMW,
          width: { size: PG.W, type: WidthType.DXA },
          children: items.map((item, i) => new Paragraph({
            children: [
              new TextRun({ text: (i+1) + ".  " + item[0] + "  ",
                font: "Arial", size: 19, bold: true, color: C.mid }),
              new TextRun({ text: item[1], font: "Arial", size: 19, color: C.muted }),
            ],
            spacing: { after: 120 },
          })),
        })],
      })],
    }),
  ];
}

// ── Parse excerpt markdown string ─────────────────────────────────────────
function parseExcerpt(excStr) {
  // Excerpt number — handle "1", "1.", "1  " etc.
  const numMatch = excStr.match(/\*\*Excerpt_Number:\*\*\s*([^\n\s]+)/);
  const num      = numMatch ? numMatch[1].replace(/\.$/, "") : "?";

  // Extract body after "Excerpt_Text:** Excerpt:" — grab everything to end, strip trailing ---
  const textMatch = excStr.match(/\*\*Excerpt_Text:\*\*\s*Excerpt:[ \t]*\n([\s\S]+)/);
  const rawBody   = textMatch
    ? textMatch[1].replace(/\n---[ \t]*$/, "").trim()
    : excStr;

  // Reference: match both **bold** and plain formats, across one or two lines
  // Pattern 1: Reference:\n**text**
  // Pattern 2: Reference:\nplain text
  // Pattern 3: Reference: plain text (inline)
  const refMatch =
    rawBody.match(/Reference:[ \t]*\n\*\*([^*\n]+)\*\*/) ||
    rawBody.match(/Reference:[ \t]*\n([^\n*][^\n]+)/)     ||
    rawBody.match(/Reference:[ \t]+([^\n]+)/);
  const ref = refMatch ? refMatch[1].trim() : "";

  // Strip the reference block from body
  const bodyStripped = rawBody
    .replace(/\nReference:[ \t]*\n\*\*[^*]+\*\*[ \t]*/g, "")
    .replace(/\nReference:[ \t]*\n[^\n]+[ \t]*/g, "")
    .replace(/\nReference:[ \t]+[^\n]+/g, "")
    .trim();

  // Section heading from ## line
  const h2Match = bodyStripped.match(/^##[ \t]+(.+)/m);
  const heading = h2Match ? h2Match[1].trim() : "";

  // Split into lines, skip top-level # headings and blank lines
  const bodyLines = bodyStripped
    .split("\n")
    .filter(l => !l.trim().startsWith("# ") && l.trim() !== "")
    .map(l => l.trim());

  return { num, ref, heading, bodyLines };
}

// ── Limitations: first line = label, rest = body ──────────────────────────
function parseLimitations(str) {
  const lines = (str || "").split("\n");
  const label = lines[0].trim();
  const body  = lines.slice(1).join("\n").replace(/^\s*\n/, "").trim();
  return { label, body };
}

// ═════════════════════════════════════════════════════════════════════════
// DOCUMENT BUILDER
// ═════════════════════════════════════════════════════════════════════════

async function buildDocument(data, meta) {
  const patentNumber  = data.Patent_Number || meta.Patent_Number || "Unknown";
  const title         = data.Title         || meta.Title         || "Patent Analysis Report";
  const owner         = data.Owner         || meta.Owner         || "";
  const standard      = data.Standard      || meta.Standard      || "";
  const claimNumber   = data.Claim_Number  || "";
  const claimText     = data.Claim         || "";
  const claimCategory = data.Claim_Category|| "";
  const pctMapped     = data.Mapped_Percentage || "";
  const pctWeighted   = data["Mapped_Percentage_(Weighted)"] || "";
  const essDecision   = data.Essentiality_Conclusion || "";
  const opinion       = data.Summary       || "";
  const mappingItems  = data.Mapping_Summary || [];
  const charts        = data.Claim_Charts  || [];
  const { label: limLabel, body: limBody } = parseLimitations(data["Limitation(s)"] || "");
  const claimLabel    = claimNumber + " \u2014 " + claimCategory + " Claim";

  // ── Section 1: Portrait — Identity, Summary, Mapping ──────────────────
  const section1Children = [
    // Title area
    new Paragraph({
      children: [
        new TextRun({ text: patentNumber + "  ", font: "Arial", size: 17, bold: true,
          color: C.orange, characterSpacing: 40 }),
        new TextRun({ text: standard + "  ", font: "Arial", size: 17, bold: true,
          color: C.navy, characterSpacing: 40 }),
        new TextRun({ text: claimNumber + " \u00B7 " + claimCategory,
          font: "Arial", size: 17, bold: true, color: C.navy, characterSpacing: 40 }),
      ],
      spacing: { before: 320, after: 120 },
    }),
    new Paragraph({
      children: [new TextRun({ text: title, font: "Georgia", size: 52, bold: true, color: C.navy })],
      spacing: { after: 280 },
    }),
    // Identity grid
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [Math.floor(PG.W/3), Math.floor(PG.W/3), PG.W - Math.floor(PG.W/3)*2],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          { label: "Patent Number", value: patentNumber },
          { label: "Owner",         value: owner },
          { label: "Standard",      value: standard },
        ].map(cell => new TableCell({
          borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
            left: noBorder, right: solidBorder(C.rule,4) },
          shading: shade(C.white), margins: CMW,
          width: { size: Math.floor(PG.W/3), type: WidthType.DXA },
          children: [
            new Paragraph({ children: [new TextRun({ text: cell.label.toUpperCase(),
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
              spacing: { after: 40 } }),
            new Paragraph({ children: [new TextRun({ text: cell.value,
              font: "Arial", size: 20, bold: true, color: C.navy })],
              spacing: { after: 0 } }),
          ],
        }))
      })],
    }),
    emptyPara(),
    claimBlock(claimLabel, claimText),
    emptyPara(),
    ...sectionHeading("Executive Summary"),
    summaryCardTable([
      { label: "Claim Number",         value: claimNumber },
      { label: "Claim Category",       value: claimCategory },
      { label: "Pct. Mapped",          value: pctMapped },
      { label: "Essentiality Decision",value: essDecision, highlight: true, small: true },
    ]),
    emptyPara(),
    summaryCardTable([
      { label: "Weighted Mapping", value: pctWeighted },
      { label: "Limitations",      value: limLabel, small: true },
    ]),
    emptyPara(),
    // Opinion box
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [PG.W],
      borders: noBorders,
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
            left: solidBorder(C.rule,4), right: solidBorder(C.rule,4) },
          shading: shade(C.white), margins: CMW,
          width: { size: PG.W, type: WidthType.DXA },
          children: [
            new Paragraph({ children: [new TextRun({ text: "Opinion",
              font: "Georgia", size: 28, bold: true, color: C.navy })],
              spacing: { after: 120 } }),
            new Paragraph({ children: [run(opinion, { size: 19, color: C.mid })],
              spacing: { after: 160 } }),
            new Paragraph({ children: [new TextRun({ text: "Limitations Detail",
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
              shading: shade(C.surfaceAlt), spacing: { after: 80, before: 120 } }),
            new Paragraph({ children: [run(limBody, { size: 19, color: C.mid })],
              spacing: { after: 0 } }),
          ],
        })],
      })],
    }),
    emptyPara(),
    ...sectionHeading("Mapping Summary"),
    ...mappingItems.flatMap((item, i) => [
      mappingItem(i + 1, item.Key_Feature, item.Conclusions, item.Brief_Rationale),
      emptyPara(),
    ]),
  ];

  // ── Section 2: Landscape — Claim Chart ────────────────────────────────
  const section2Children = [
    new Paragraph({
      children: [new TextRun({ text: "Claim Chart", font: "Georgia",
        size: 32, bold: true, color: C.navy })],
      spacing: { before: 480, after: 160 },
      border: { bottom: solidBorder(C.rule, 4) },
    }),
    ...charts.flatMap(chart => {
      const feat = chart.Claim_Feature || {};
      const dec  = chart.Decision      || {};
      const ana  = chart.Analysis      || {};
      const excRaw = chart.Cited_Excerpts || [];

      const colW   = Math.floor(PGL.W / 2);
      const innerW = colW - (CMW.left + CMW.right);

      const analysisChildren = [
        ...analysisParagraphs(
          ana.Interpretation  || "",
          ana.Mapping_Summary || "",
          ana.Differences     || "",
          ana.Overall_Opinion || ""
        ),
        justificationPanel(dec.Justification || "", innerW),
      ];

      const excerptTables = excRaw.map((excStr, i) => {
        const exc = parseExcerpt(excStr);
        return excerptItem(exc.num, exc.ref, exc.heading, exc.bodyLines, innerW);
      });

      return featureBlock(
        feat.Index || (charts.indexOf(chart) + 1),
        feat.Text  || "",
        dec.Disclosure || "",
        dec.Essentiality_Classification || "",
        analysisChildren,
        excerptTables,
        PGL.W
      );
    }),
  ];

  // ── Section 3: Portrait — Disclaimer ──────────────────────────────────
  const section3Children = [...disclaimerSection(), emptyPara()];

  const doc = new Document({
    sections: [
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        headers: { default: makeHeader(PG.W) },
        footers: { default: makeFooter() },
        children: section1Children,
      },
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        headers: { default: makeHeader(PGL.W) },
        footers: { default: makeFooter() },
        children: section2Children,
      },
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        headers: { default: makeHeader(PG.W) },
        footers: { default: makeFooter() },
        children: section3Children,
      },
    ],
  });

  return Packer.toBuffer(doc);
}

// ═════════════════════════════════════════════════════════════════════════
// EXPRESS ROUTES
// ═════════════════════════════════════════════════════════════════════════

// Health check
app.get("/", (req, res) => {
  res.json({ status: "ok", service: "ipmind-docx-service" });
});

// Main endpoint — POST the IPMIND analysis JSON
// Optional query params: patent, title, owner, standard (override META)
app.post("/generate", async (req, res) => {
  try {
    let body = req.body;

    // Accept both array [ {...} ] and plain object { ... }
    const data = Array.isArray(body) ? body[0] : body;

    // Patent metadata can also be passed as query params
    const meta = {
      Patent_Number: req.query.patent  || "",
      Title:         req.query.title   || "",
      Owner:         req.query.owner   || "",
      Standard:      req.query.standard|| "",
    };

    const buf      = await buildDocument(data, meta);
    const safeName = (data.Patent_Number || meta.Patent_Number || "report")
      .replace(/[^A-Za-z0-9_-]/g, "_");
    const filename = safeName + "_report.docx";

    res.setHeader("Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="' + filename + '"');
    res.setHeader("Content-Length", buf.length);
    res.send(buf);

  } catch (err) {
    console.error("Error generating docx:", err);
    res.status(500).json({ error: err.message });
  }
});

// ═════════════════════════════════════════════════════════════════════════
// START
// ═════════════════════════════════════════════════════════════════════════

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("IPMIND docx service running on port " + PORT));
