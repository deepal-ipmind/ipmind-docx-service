"use strict";
const express = require("express");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber,
  TabStopType, TabStopPosition, PageBreak, PageOrientation, SectionType
} = require("docx");

const app  = express();
app.use(express.json({ limit: "10mb" }));

// ── API Key Authentication ────────────────────────────────────────────────
const API_KEY = process.env.IPMIND_API_KEY;

function requireApiKey(req, res, next) {
  if (req.path === "/") return next();
  if (!API_KEY) {
    return res.status(503).json({ error: "Service misconfigured: IPMIND_API_KEY is not set." });
  }
  const provided = req.headers["x-api-key"];
  if (!provided) {
    return res.status(401).json({ error: "Missing x-api-key header." });
  }
  if (provided !== API_KEY) {
    return res.status(403).json({ error: "Invalid API key." });
  }
  next();
}

app.use(requireApiKey);

// ── Embedded logo (base64 PNG, navy background) ──────────────────────────
const LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAUAAAABpCAYAAABLRPrgAAAABmJLR0QA/wD/AP+gvaeTAAAY4klEqVR4nO2bebQdVbGHf8UkyEwMiRAiCkQiyDwKKAJRkEFRRuGpCD4BEVBQUViCPBFkkMElLMUBUZmePomggEBkkEEgAQXDPIYkkBBISEiAJPd7f+x9vCcnfXo6Q5+bW99aWRfO6V27uvt09d69a1dJjuM4juM4juM4juM4juM4juM4juM4juM4jjPgsaoVKAvwTknvkTRM0rskrSxp2bpD5kh6XdJ0SS9Kmmpmfd3W03Gc3mVAGEBgbUk7SNpa0uaSRktaq6CYtyQ9KelhSeMl/UPS/Wb2VhtVdRxnANGTBhBYWtJOkj4paQ9J7044bI6k5yVNkfSKpNmS5kuaJ2kVSctIWkPSUEkjFQzm0g0y5kq6Q9L1kq41s8ntPhfHcXqXnjKAwAckHSHpYEnD674aIelOSfdKmiDpETObWlD2cpJGSdpU0lYKI8ot1G8U+2Ifl0m6xszmlj4Rx3GcPAAG7A2MY1H+BZwGbAks1aG+VwcOAK4AXq/r+zXgHGBEJ/p1HGeQEw3fvsA/GwzP+cAHK9BnBeDgBkP8FnCJG0LHcdoGsB1wT52heQo4Mq7sVg6wMfAr4O2o30+BJ4Glq9bNcZweB1gu+vVqo75bgPdWrVdRCBmxv1y30v04sFnVejmOMwAgbGF7oW7kdBoZ6ex7BWBL4O6oex/wU2ClqvVyHGcAAawCXAwsjMbkRULNjp5cPCAURvpNnb5PALtWrZfjOAOYOKKqz9T8LPDVXhlVEfIN/paQxoo47f0WHuLiOE67APYGJtQZwpmEAOqurxgDKxKKNd1ep8+c6L8c2m19HMcZBBBq/e4B/JVF64Q8Sgg12Z6StUdy9D0UOAS4hkXLYL4EnAokFVRyHGcJoadi14BRkr4o6RAtWpltlqQ7Jd0jabykiZImmRkFZK8gaUNJm0jaWtJOkj6o/mswXyFrzGWSxnrlNsdZ8ukpA1iDECu4naRPSdpd0sZaXNc3JT2rkOJquqTXas3jsctLWl3SUEkjJa2V0NVMSeMUUmz9ycxmtPVEHMfpaXrSADYCrClpB4U8gJtJGi1pHRXTf7akxyU9LGmCpLsl/dPMFrZXW8dxBgoDwgAmEVdkR0oaLmmIpJUk1eIKF0p6XWHq/LKkyWb2ShV6Oo7jOI7jOI7jOI7jOI7jOI7jOI7jOI7jOI7jOE4H+X82HH2b5xE8bQAAAABJRU5ErkJggg==";
const logoData = Buffer.from(LOGO_B64, "base64");

// ── Palette ──────────────────────────────────────────────────────────────
const C = {
  orange:     "FF6734", navy:      "0F1F38", ink:       "1C1C2E",
  mid:        "4A4A6A", muted:     "7A7A96", rule:      "E2E2ED",
  surfaceAlt: "F4F4F0", amberText: "8A5A00", amberBg:   "FDF5E0",
  greenText:  "1A6B4A", greenBg:   "EAF5EF", white:     "FFFFFF",
};

// ── Page geometry ─────────────────────────────────────────────────────────
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

// ── Safe string helper ────────────────────────────────────────────────────
function safeStr(val) {
  if (val === null || val === undefined) return "";
  // eslint-disable-next-line no-control-regex
  return String(val).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/g, "");
}

// ── Text helpers ──────────────────────────────────────────────────────────
function run(text, opts = {}) {
  return new TextRun({ text: safeStr(text), font: "Arial", size: 20, color: C.ink, ...opts });
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
              children: [new TextRun({ text: safeStr(label).toUpperCase(),
                font: "Arial", size: 16, bold: true, color: C.orange, characterSpacing: 40 })],
              spacing: { after: 80 },
            }),
            new Paragraph({
              children: [new TextRun({ text: safeStr(text), font: "Arial", size: 19, color: C.ink, italics: true })],
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
            children: [new TextRun({ text: safeStr(c.label).toUpperCase(),
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
            spacing: { after: 60 },
          }),
          new Paragraph({
            children: [new TextRun({ text: safeStr(c.value), font: "Arial",
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
              children: [new TextRun({ text: safeStr(featureText), font: "Arial",
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
                  new TextRun({ text: safeStr(conclusion), font: "Arial", size: 19, color: C.mid }),
                ],
                spacing: { after: 80 },
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: "Brief Rationale:  ", font: "Arial", size: 19, bold: true, color: C.ink }),
                  new TextRun({ text: safeStr(rationale), font: "Arial", size: 19, color: C.mid }),
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
              children: [new TextRun({ text: safeStr(text), font: "Arial", size: 19, color: C.mid })],
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
  const lines = Array.isArray(mappingDetail)
    ? mappingDetail
    : String(safeStr(mappingDetail) || "").split(/\n\n+/).filter(Boolean);

  return [
    subHeading("Interpretation"),
    new Paragraph({ children: [run(safeStr(interpretation), { size: 19, color: C.mid })], spacing: { after: 120 } }),
    subHeading("Mapping Summary"),
    ...lines.map(line => new Paragraph({
      children: [run(safeStr(line).replace(/\*\*/g, ""), { size: 19, color: C.mid })],
      spacing: { after: 80 },
    })),
    subHeading("Differences"),
    new Paragraph({ children: [run(safeStr(differences), { size: 19, color: C.mid })], spacing: { after: 120 } }),
    subHeading("Overall Opinion"),
    new Paragraph({ children: [run(safeStr(opinion), { size: 19, color: C.mid })], spacing: { after: 160 } }),
  ];
}

// ── Excerpt item ──────────────────────────────────────────────────────────
function excerptItem(num, ref, heading, bodyLines, W = PG.W) {
  const labelW = 1100;
  const refW   = W - labelW;
  const bodyChildren = [];
  if (heading) {
    bodyChildren.push(new Paragraph({
      children: [new TextRun({ text: safeStr(heading), font: "Arial", size: 16,
        bold: true, color: C.navy, characterSpacing: 30 })],
      spacing: { after: 80 },
    }));
  }
  bodyLines.forEach(line => {
    bodyChildren.push(new Paragraph({
      children: [new TextRun({ text: safeStr(line), font: "Courier New", size: 15, color: C.mid })],
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
              children: [new TextRun({ text: safeStr(ref), font: "Courier New", size: 15, color: C.muted })],
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
  const verdictText = safeStr(disclosure) + "  ·  " + safeStr(essentiality);
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
              children: [new TextRun({ text: safeStr(num), font: "Arial",
                size: 28, bold: true, color: C.white })],
              alignment: AlignmentType.CENTER, spacing: { after: 0 },
            })],
          }),
          new TableCell({
            borders: noBorders, shading: shade(C.navy),
            margins: { top: 120, bottom: 120, left: 80, right: 160 },
            width: { size: W - 400, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              children: [new TextRun({ text: safeStr(featureText), font: "Arial",
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

// ── Methodology guide — Word ──────────────────────────────────────────────
// Renders a "Key to Terms" reference table at the top of the landscape
// claim chart section. Silently omitted when data.Methodology is absent.
function methodologyDocx(methodology, W = PGL.W) {
  if (!methodology) return [];

  const metrics    = methodology.universal_metrics       || {};
  const disclosure = methodology.disclosure_categories   || [];
  const ess        = methodology.essentiality_tiers       || [];

  const labelW = 2800;
  const defW   = W - labelW;

  // Group header: full-width navy bar with white uppercase label
  function groupHeaderRow(title) {
    return new TableRow({
      children: [new TableCell({
        columnSpan: 2,
        borders: {
          top: solidBorder(C.rule, 4), bottom: solidBorder(C.rule, 4),
          left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4),
        },
        shading: shade(C.navy),
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        width: { size: W, type: WidthType.DXA },
        children: [new Paragraph({
          children: [new TextRun({ text: title.toUpperCase(),
            font: "Arial", size: 15, bold: true, color: C.white, characterSpacing: 60 })],
          spacing: { after: 0 },
        })],
      })],
    });
  }

  // Term row: label cell (left) | definition cell (right)
  function termRow(label, definition, labelColor = C.navy, labelBg = C.surfaceAlt) {
    return new TableRow({
      children: [
        new TableCell({
          borders: {
            top: noBorder, bottom: solidBorder(C.rule, 4),
            left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4),
          },
          shading: shade(labelBg),
          margins: CM,
          width: { size: labelW, type: WidthType.DXA },
          verticalAlign: VerticalAlign.TOP,
          children: [new Paragraph({
            children: [new TextRun({ text: safeStr(label),
              font: "Arial", size: 17, bold: true, color: labelColor })],
            spacing: { after: 0 },
          })],
        }),
        new TableCell({
          borders: {
            top: noBorder, bottom: solidBorder(C.rule, 4),
            left: noBorder, right: solidBorder(C.rule, 4),
          },
          shading: shade(C.white),
          margins: CM,
          width: { size: defW, type: WidthType.DXA },
          verticalAlign: VerticalAlign.TOP,
          children: [new Paragraph({
            children: [new TextRun({ text: safeStr(definition),
              font: "Arial", size: 17, color: C.mid })],
            spacing: { after: 0 },
          })],
        }),
      ],
    });
  }

  // Pick label colour/bg based on category type
  function disclosureStyle(label) {
    const l = (label || "").toLowerCase();
    if (l.includes("not disclosed"))                         return { color: "8A0000", bg: "FDF0F0" };
    if (l.includes("explicitly") || l.includes("implied"))  return { color: C.greenText, bg: C.greenBg };
    return { color: C.amberText, bg: C.amberBg }; // partial, functional equivalence
  }

  function essStyle(label) {
    const l = (label || "").toLowerCase();
    if (l.includes("not essential") || l.includes("non-technical")) return { color: "8A0000", bg: "FDF0F0" };
    if (l.includes("conditional"))   return { color: C.amberText, bg: C.amberBg };
    if (l.includes("essential"))     return { color: C.greenText,  bg: C.greenBg };
    return { color: C.navy, bg: C.surfaceAlt }; // implementation matter etc.
  }

  const rows = [];

  // Metrics section
  rows.push(groupHeaderRow("Metrics"));
  if (metrics.percentage_mapped) rows.push(termRow("Percentage Mapped", metrics.percentage_mapped));
  if (metrics.weighted_mapping)  rows.push(termRow("Weighted Mapping",  metrics.weighted_mapping));

  // Disclosure categories section
  if (disclosure.length > 0) {
    rows.push(groupHeaderRow("Disclosure Categories"));
    disclosure.forEach(item => {
      const s = disclosureStyle(item.label);
      rows.push(termRow(safeStr(item.label), safeStr(item.definition), s.color, s.bg));
    });
  }

  // Essentiality tiers section
  if (ess.length > 0) {
    rows.push(groupHeaderRow("Essentiality Tiers"));
    ess.forEach(item => {
      const s = essStyle(item.label);
      rows.push(termRow(safeStr(item.label), safeStr(item.definition), s.color, s.bg));
    });
  }

  return [
    new Paragraph({
      children: [new TextRun({ text: "Key to Terms", font: "Georgia",
        size: 28, bold: true, color: C.navy })],
      spacing: { before: 320, after: 160 },
      border: { bottom: solidBorder(C.rule, 4) },
    }),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [labelW, defW],
      borders: noBorders,
      rows,
    }),
    emptyPara(),
    emptyPara(),
  ];
}

// ── Parse excerpt markdown string ─────────────────────────────────────────
function parseExcerpt(excStr) {
  const numMatch = excStr.match(/\*\*Excerpt_Number:\*\*\s*([^\n\s]+)/);
  const num      = numMatch ? numMatch[1].replace(/\.$/, "") : "?";
  const textMatch = excStr.match(/\*\*Excerpt_Text:\*\*\s*Excerpt:[ \t]*\n([\s\S]+)/);
  const rawBody   = textMatch
    ? textMatch[1].replace(/\n---[ \t]*$/, "").trim()
    : excStr;
  const refMatch =
    rawBody.match(/Reference:[ \t]*\n\*\*([^*\n]+)\*\*/) ||
    rawBody.match(/Reference:[ \t]*\n([^\n*][^\n]+)/)     ||
    rawBody.match(/Reference:[ \t]+([^\n]+)/);
  const ref = refMatch ? refMatch[1].trim() : "";
  const bodyStripped = rawBody
    .replace(/\nReference:[ \t]*\n\*\*[^*]+\*\*[ \t]*/g, "")
    .replace(/\nReference:[ \t]*\n[^\n]+[ \t]*/g, "")
    .replace(/\nReference:[ \t]+[^\n]+/g, "")
    .trim();
  const h2Match = bodyStripped.match(/^##[ \t]+(.+)/m);
  const heading = h2Match ? h2Match[1].trim() : "";
  const bodyLines = bodyStripped
    .split("\n")
    .filter(l => !l.trim().startsWith("# ") && l.trim() !== "")
    .map(l => l.trim());
  return { num, ref, heading, bodyLines };
}

// ── Limitations parser ────────────────────────────────────────────────────
function parseLimitations(str) {
  const lines = (str || "").split("\n");
  const label = lines[0].trim();
  const body  = lines.slice(1).join("\n").replace(/^\s*\n/, "").trim();
  return { label, body };
}

// ═════════════════════════════════════════════════════════════════════════
// RESTRICTED USE NOTICE
// ═════════════════════════════════════════════════════════════════════════

const RESTRICTED_NOTICE_TEXT =
  "This report is confidential and provided solely for internal use in connection " +
  "with patent licensing, portfolio evaluation, or standards-related strategy. It must " +
  "not be published, posted, or circulated to any third party without IP Mind\u2019s prior " +
  "written consent. Where disclosure to a counterparty is necessary, the report may be " +
  "shared in full or in part provided the counterparty is bound by a written " +
  "confidentiality undertaking that places equivalent restrictions on use and further " +
  "distribution, and that requires attribution of IP Mind\u2019s authorship to be retained. " +
  "The recipient must not use this report to replicate, benchmark, or train models " +
  "intended to reproduce IP Mind\u2019s methodology or outputs, or to develop competing " +
  "analysis products or services.";

function restrictedNoticePage() {
  return [
    new Paragraph({
      children: [new TextRun({ text: "RESTRICTED USE NOTICE", font: "Arial",
        size: 17, bold: true, color: C.orange, characterSpacing: 80 })],
      spacing: { before: 480, after: 240 },
      border: { bottom: solidBorder(C.orange, 8) },
    }),
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [60, PG.W - 60],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: { top: noBorder, bottom: noBorder,
              left: solidBorder(C.orange, 12), right: noBorder },
            shading: shade(C.amberBg),
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            width: { size: 60, type: WidthType.DXA },
            children: [emptyPara()],
          }),
          new TableCell({
            borders: noBorders,
            shading: shade(C.amberBg),
            margins: { top: 200, bottom: 200, left: 240, right: 240 },
            width: { size: PG.W - 60, type: WidthType.DXA },
            children: [
              new Paragraph({
                children: [new TextRun({ text: RESTRICTED_NOTICE_TEXT,
                  font: "Arial", size: 19, color: C.amberText })],
                spacing: { after: 0 },
              }),
            ],
          }),
        ],
      })],
    }),
    emptyPara(),
  ];
}

// ═════════════════════════════════════════════════════════════════════════
// DOCUMENT BUILDER
// ═════════════════════════════════════════════════════════════════════════

async function buildDocument(data, meta, restricted) {
  const patentNumber  = safeStr(data.Patent_Number || meta.Patent_Number || "Unknown");
  const title         = safeStr(data.Title         || meta.Title         || "Patent Analysis Report");
  const owner         = safeStr(data.Owner         || meta.Owner         || "");
  const standard      = safeStr(data.Standard      || meta.Standard      || "");
  const claimNumber   = safeStr(data.Claim_Number  || "");
  const claimText     = safeStr(data.Claim         || "");
  const claimCategory = safeStr(data.Claim_Category|| "");
  const pctMapped     = safeStr(data.Mapped_Percentage || "");
  const pctWeighted   = safeStr(data["Mapped_Percentage_(Weighted)"] || "");
  const essDecision   = safeStr(data.Essentiality_Conclusion || "");
  const opinion       = safeStr(data.Summary       || "");
  const methodology   = data.Methodology || null;
  const mappingItems  = (data.Mapping_Summary || []).slice().sort((a, b) => {
    return (parseInt(a.Index, 10) || 0) - (parseInt(b.Index, 10) || 0);
  });
  const charts        = (data.Claim_Charts || []).slice().sort((a, b) => {
    const ai = parseInt((a.Claim_Feature || {}).Index, 10) || 0;
    const bi = parseInt((b.Claim_Feature || {}).Index, 10) || 0;
    return ai - bi;
  });
  const { label: limLabelRaw, body: limBodyRaw } = parseLimitations(data["Limitation(s)"] || "");
  const limLabel = safeStr(limLabelRaw);
  const limBody  = safeStr(limBodyRaw);
  const claimLabel = safeStr(claimNumber) + " \u2014 " + safeStr(claimCategory) + " Claim";

  const section1Children = [
    new Paragraph({
      children: [
        new TextRun({ text: safeStr(patentNumber) + "  ", font: "Arial", size: 17, bold: true,
          color: C.orange, characterSpacing: 40 }),
        new TextRun({ text: safeStr(standard) + "  ", font: "Arial", size: 17, bold: true,
          color: C.navy, characterSpacing: 40 }),
        new TextRun({ text: safeStr(claimNumber) + " \u00B7 " + safeStr(claimCategory),
          font: "Arial", size: 17, bold: true, color: C.navy, characterSpacing: 40 }),
      ],
      spacing: { before: 320, after: 120 },
    }),
    new Paragraph({
      children: [new TextRun({ text: safeStr(title), font: "Georgia", size: 52, bold: true, color: C.navy })],
      spacing: { after: 280 },
    }),
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [Math.floor(PG.W/3), Math.floor(PG.W/3), PG.W - Math.floor(PG.W/3)*2],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          { label: "Patent Number", value: safeStr(patentNumber) },
          { label: "Owner",         value: safeStr(owner) },
          { label: "Standard",      value: safeStr(standard) },
        ].map(cell => new TableCell({
          borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
            left: noBorder, right: solidBorder(C.rule,4) },
          shading: shade(C.white), margins: CMW,
          width: { size: Math.floor(PG.W/3), type: WidthType.DXA },
          children: [
            new Paragraph({ children: [new TextRun({ text: safeStr(cell.label).toUpperCase(),
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
              spacing: { after: 40 } }),
            new Paragraph({ children: [new TextRun({ text: safeStr(cell.value),
              font: "Arial", size: 20, bold: true, color: C.navy })],
              spacing: { after: 0 } }),
          ],
        }))
      })],
    }),
    emptyPara(),
    ...(restricted ? [new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [PG.W],
      borders: noBorders,
      rows: [new TableRow({
        children: [new TableCell({
          borders: noBorders,
          shading: shade(C.amberBg),
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          width: { size: PG.W, type: WidthType.DXA },
          children: [new Paragraph({
            children: [
              new TextRun({ text: "\u26A0\uFE0F  RESTRICTED USE", font: "Arial",
                size: 16, bold: true, color: C.amberText, characterSpacing: 60 }),
              new TextRun({ text: "   \u00B7   This report is subject to restricted use. See notice at end of document.",
                font: "Arial", size: 16, color: C.amberText }),
            ],
            spacing: { after: 0 },
          })],
        })],
      })],
    }), emptyPara()] : []),
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
      mappingItem(i + 1, safeStr(item.Key_Feature), safeStr(item.Conclusions), safeStr(item.Brief_Rationale)),
      emptyPara(),
    ]),
  ];

  const section2Children = [
    new Paragraph({
      children: [new TextRun({ text: "Claim Chart", font: "Georgia",
        size: 32, bold: true, color: C.navy })],
      spacing: { before: 480, after: 160 },
      border: { bottom: solidBorder(C.rule, 4) },
    }),
    // ── Key to Terms — only rendered when Methodology is present in payload
    ...methodologyDocx(methodology, PGL.W),
    ...charts.flatMap(chart => {
      const feat = chart.Claim_Feature || {};
      const dec  = chart.Decision      || {};
      const ana  = chart.Analysis      || {};
      const excRaw = chart.Cited_Excerpts || [];
      const colW   = Math.floor(PGL.W / 2);
      const innerW = colW - (CMW.left + CMW.right);
      const analysisChildren = [
        ...analysisParagraphs(
          safeStr(ana.Interpretation  || ""),
          ana.Mapping_Summary || "",
          safeStr(ana.Differences     || ""),
          safeStr(ana.Overall_Opinion || "")
        ),
        justificationPanel(safeStr(dec.Justification || ""), innerW),
      ];
      const excerptTables = excRaw.map((excStr) => {
        const exc = parseExcerpt(excStr);
        return excerptItem(exc.num, exc.ref, exc.heading, exc.bodyLines, innerW);
      });
      return featureBlock(
        feat.Index || (charts.indexOf(chart) + 1),
        safeStr(feat.Text  || ""),
        safeStr(dec.Disclosure || ""),
        safeStr(dec.Essentiality_Classification || ""),
        analysisChildren,
        excerptTables,
        PGL.W
      );
    }),
  ];

  const section3Children = [
    ...(restricted ? restrictedNoticePage() : []),
    ...disclaimerSection(),
    emptyPara(),
  ];

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
// EXPRESS ROUTES — /generate
// ═════════════════════════════════════════════════════════════════════════

app.get("/", (req, res) => {
  res.json({ status: "ok", service: "ipmind-docx-service" });
});

app.post("/generate", async (req, res) => {
  try {
    const body = req.body;
    const data = Array.isArray(body) ? body[0] : body;
    const meta = {
      Patent_Number: req.query.patent   || "",
      Title:         req.query.title    || "",
      Owner:         req.query.owner    || "",
      Standard:      req.query.standard || "",
    };
    const restricted = req.query.restricted === "true" || data.Restricted_Use === true;
    const buf      = await buildDocument(data, meta, restricted);
    const safeName = (data.Patent_Number || meta.Patent_Number || "report")
      .replace(/[^A-Za-z0-9_-]/g, "_");
    res.setHeader("Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="' + safeName + '_report.docx"');
    res.setHeader("Content-Length", buf.length);
    res.send(buf);
  } catch (err) {
    console.error("Error generating docx:", err);
    res.status(500).json({ error: err.message });
  }
});

// ═════════════════════════════════════════════════════════════════════════
// HTML BUILDER
// ═════════════════════════════════════════════════════════════════════════

const RESTRICTED_NOTICE_HTML =
  "This report is confidential and provided solely for internal use in connection " +
  "with patent licensing, portfolio evaluation, or standards-related strategy. It must " +
  "not be published, posted, or circulated to any third party without IP Mind\u2019s prior " +
  "written consent. Where disclosure to a counterparty is necessary, the report may be " +
  "shared in full or in part provided the counterparty is bound by a written " +
  "confidentiality undertaking that places equivalent restrictions on use and further " +
  "distribution, and that requires attribution of IP Mind\u2019s authorship to be retained. " +
  "The recipient must not use this report to replicate, benchmark, or train models " +
  "intended to reproduce IP Mind\u2019s methodology or outputs, or to develop competing " +
  "analysis products or services.";

function buildHtml(data, meta, restricted) {
  const patentNumber  = safeStr(data.Patent_Number  || meta.Patent_Number  || "");
  const title         = safeStr(data.Title          || meta.Title          || "");
  const owner         = safeStr(data.Owner          || meta.Owner          || "");
  const standard      = safeStr(data.Standard       || meta.Standard       || "");
  const claimNumber   = safeStr(data.Claim_Number   || "");
  const claimText     = safeStr(data.Claim          || "");
  const claimCategory = safeStr(data.Claim_Category || "");
  const pctMapped     = safeStr(data.Mapped_Percentage || "");
  const pctWeighted   = safeStr(data["Mapped_Percentage_(Weighted)"] || "");
  const essDecision   = safeStr(data.Essentiality_Conclusion || "");
  const opinion       = safeStr(data.Summary        || "");
  const methodology   = data.Methodology || null;
  const mappingItems  = (data.Mapping_Summary || []).slice().sort((a, b) => {
    return (parseInt(a.Index, 10) || 0) - (parseInt(b.Index, 10) || 0);
  });
  const charts        = (data.Claim_Charts || []).slice().sort((a, b) => {
    const ai = parseInt((a.Claim_Feature || {}).Index, 10) || 0;
    const bi = parseInt((b.Claim_Feature || {}).Index, 10) || 0;
    return ai - bi;
  });
  const { label: limLabelRaw, body: limBodyRaw } = parseLimitations(data["Limitation(s)"] || "");
  const limLabel = safeStr(limLabelRaw);
  const limBody  = safeStr(limBodyRaw);
  const claimLabel = `${claimNumber} \u2014 ${claimCategory} Claim`;

  function esc(str) {
    return safeStr(str)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function renderInline(str) {
    return safeStr(str)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
      .replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, "<em>$1</em>");
  }

  function renderMdTable(tableLines) {
    let thead = "", tbody = "";
    tableLines.forEach((line, idx) => {
      if (/^\|[\s\-:|]+\|$/.test(line.trim())) return;
      const cells = line.split("|").slice(1, -1);
      if (idx === 0) {
        thead = "<thead><tr>" + cells.map(c => `<th>${renderInline(c.trim())}</th>`).join("") + "</tr></thead>";
      } else {
        tbody += "<tr>" + cells.map(c => `<td>${renderInline(c.trim())}</td>`).join("") + "</tr>";
      }
    });
    return `<table class="exc-table">${thead}<tbody>${tbody}</tbody></table>`;
  }

  function renderMdBlock(md) {
    if (!md) return "";
    const lines = md.split("\n");
    let out = "", i = 0;
    while (i < lines.length) {
      const line = lines[i];
      const trimmed = line.trim();
      if (trimmed.startsWith("|")) {
        const tableLines = [];
        while (i < lines.length && lines[i].trim().startsWith("|")) { tableLines.push(lines[i]); i++; }
        out += renderMdTable(tableLines); continue;
      }
      if (trimmed.startsWith("## ")) { out += `<p class="exc-subhead">${renderInline(trimmed.slice(3))}</p>`; i++; continue; }
      if (trimmed.startsWith("# "))  { i++; continue; }
      if (trimmed.startsWith("\u2014")) {
        out += '<ul class="exc-list">';
        while (i < lines.length && lines[i].trim().startsWith("\u2014")) {
          const text = lines[i].trim().replace(/^[\u2014]\s*/, "");
          out += `<li>${renderInline(text)}</li>`; i++;
        }
        out += "</ul>"; continue;
      }
      if (/^ {2,}/.test(line) && trimmed !== "") { out += `<p class="exc-indent">${renderInline(trimmed)}</p>`; i++; continue; }
      if (trimmed === "") { i++; continue; }
      out += `<p>${renderInline(trimmed)}</p>`; i++;
    }
    return out;
  }

  function renderAnalysisBlock(str) {
    if (!str) return "<p></p>";
    return str.split(/\n\n+/).map(p => {
      p = p.trim();
      if (!p) return "";
      const subLines = p.split("\n").filter(l => l.trim());
      if (subLines.length > 1) return subLines.map(l => `<p>${renderInline(l.trim())}</p>`).join("");
      return `<p>${renderInline(p)}</p>`;
    }).join("");
  }

  // ── Colour helpers ────────────────────────────────────────────────────
  function essClasses(decision) {
    const d = (decision || "").toLowerCase();
    if (d.includes("not essential"))  return { card: "", value: "red", dot: "dot-red", verdict: "red", badge: "badge-red" };
    if (d.includes("conditional"))    return { card: "highlight", value: "amber", dot: "dot-amber", verdict: "amber", badge: "badge-amber" };
    if (d.includes("essential"))      return { card: "highlight-green", value: "green", dot: "dot-green", verdict: "green", badge: "badge-green" };
    return { card: "highlight", value: "amber", dot: "dot-amber", verdict: "amber", badge: "badge-amber" };
  }

  function disclosureClasses(disclosure) {
    const d = (disclosure || "").toLowerCase();
    if (d.includes("not disclosed"))  return { dot: "dot-red",   verdict: "red",   badge: "badge-red"   };
    if (d.includes("explicitly") || d.includes("implied")) return { dot: "dot-green", verdict: "green", badge: "badge-green" };
    return { dot: "dot-amber", verdict: "amber", badge: "badge-amber" };
  }

  // Label CSS class for the methodology guide rows
  function disclosureLabelClass(label) {
    const l = (label || "").toLowerCase();
    if (l.includes("not disclosed"))                        return "meth-label-red";
    if (l.includes("explicitly") || l.includes("implied")) return "meth-label-green";
    return "meth-label-amber"; // partial, functional equivalence
  }

  function essLabelClass(label) {
    const l = (label || "").toLowerCase();
    if (l.includes("not essential") || l.includes("non-technical")) return "meth-label-red";
    if (l.includes("conditional"))   return "meth-label-amber";
    if (l.includes("essential"))     return "meth-label-green";
    return "meth-label-navy"; // implementation matter etc.
  }

  // ── Methodology guide HTML ────────────────────────────────────────────
  // Collapsed by default; expands on click; always printed expanded.
  function buildMethodologyHtml(meth) {
    if (!meth) return "";
    const metrics    = meth.universal_metrics       || {};
    const disclosure = meth.disclosure_categories   || [];
    const ess        = meth.essentiality_tiers       || [];

    function groupHeader(title) {
      return `<div class="meth-group-header">${esc(title)}</div>`;
    }

    function termRow(label, definition, labelClass) {
      return `
        <div class="meth-row">
          <div class="meth-label ${labelClass}">${esc(label)}</div>
          <div class="meth-def">${esc(definition)}</div>
        </div>`;
    }

    let inner = "";

    inner += groupHeader("Metrics");
    if (metrics.percentage_mapped) inner += termRow("Percentage Mapped", metrics.percentage_mapped, "meth-label-navy");
    if (metrics.weighted_mapping)  inner += termRow("Weighted Mapping",  metrics.weighted_mapping,  "meth-label-navy");

    if (disclosure.length > 0) {
      inner += groupHeader("Disclosure Categories");
      disclosure.forEach(item => {
        inner += termRow(item.label, item.definition, disclosureLabelClass(item.label));
      });
    }

    if (ess.length > 0) {
      inner += groupHeader("Essentiality Tiers");
      ess.forEach(item => {
        inner += termRow(item.label, item.definition, essLabelClass(item.label));
      });
    }

    return `
    <div class="methodology-panel">
      <button class="methodology-toggle" onclick="toggleMethodology(this)" aria-expanded="false">
        <span class="meth-toggle-left">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="flex-shrink:0"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          <span>Key to Terms</span>
        </span>
        <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
      </button>
      <div class="methodology-body">${inner}</div>
    </div>`;
  }

  function parseExcerptHtml(excStr) {
    const numMatch  = excStr.match(/\*\*Excerpt_Number:\*\*\s*([^\n\s]+)/);
    const num       = numMatch ? numMatch[1] : "?";
    const textMatch = excStr.match(/\*\*Excerpt_Text:\*\*\s*Excerpt:[ \t]*\n([\s\S]+)/);
    const rawBody   = textMatch ? textMatch[1].replace(/\n---[ \t]*$/, "").trim() : excStr;
    const refMatch  =
      rawBody.match(/Reference:[ \t]*\n\*\*([^*\n]+)\*\*/) ||
      rawBody.match(/Reference:[ \t]*\n([^\n*][^\n]+)/)     ||
      rawBody.match(/Reference:[ \t]+([^\n]+)/);
    const ref = refMatch ? refMatch[1].trim() : "";
    const bodyStripped = rawBody
      .replace(/\nReference:[ \t]*\n\*\*[^*]+\*\*[ \t]*/g, "")
      .replace(/\nReference:[ \t]*\n[^\n]+[ \t]*/g, "")
      .replace(/\nReference:[ \t]+[^\n]+/g, "")
      .trim();
    const h2Match = bodyStripped.match(/^##[ \t]+(.+)/m);
    const heading = h2Match ? h2Match[1].trim() : "";
    const bodyHtml = renderMdBlock(bodyStripped);
    return { num, ref, heading, bodyHtml };
  }

  const ec = essClasses(essDecision);

  function buildMappingSummaryHtml(items) {
    return items.map(item => {
      const conclusionsStr = item.Conclusions || "";
      const disclosurePart = conclusionsStr.includes("|")
        ? conclusionsStr.split("|")[0].trim()
        : conclusionsStr;
      const bc = disclosureClasses(disclosurePart).badge;
      return `
        <div class="mapping-item">
          <div class="mapping-item-header">
            <div class="feat-num">${esc(item.Index)}</div>
            <div class="feat-text">${esc(item.Key_Feature)}</div>
            <div><span class="badge ${bc}">${esc(disclosurePart)}</span></div>
          </div>
          <div class="mapping-item-body">
            <p><strong>Conclusion:</strong> ${esc(conclusionsStr)}</p>
            <p><strong>Brief Rationale:</strong> ${esc(item.Brief_Rationale)}</p>
          </div>
        </div>`;
    }).join("\n");
  }

  function buildClaimChartHtml(charts) {
    return charts.map(chart => {
      const feat  = chart.Claim_Feature || {};
      const dec   = chart.Decision      || {};
      const ana   = chart.Analysis      || {};
      const excRaw = chart.Cited_Excerpts || [];
      const disclosure    = dec.Disclosure || "";
      const essClass      = dec.Essentiality_Classification || "";
      const justification = dec.Justification || "";
      const dc = disclosureClasses(disclosure);
      const fc = essClasses(essClass);
      const parsedExcs = excRaw.map(parseExcerptHtml);
      const excItemsHtml = parsedExcs.map(exc => `
              <div class="excerpt-item">
                <div class="excerpt-item-header">
                  <span class="exc-num">Excerpt ${esc(exc.num)}</span>
                  <span class="exc-ref">${esc(exc.ref)}</span>
                </div>
                <div class="excerpt-item-body">
                  ${exc.heading ? `<h4>${esc(exc.heading)}</h4>` : ""}
                  ${exc.bodyHtml}
                </div>
              </div>`).join("\n");
      return `
      <div class="claim-feature-block">
        <div class="cfb-header">
          <div class="feat-num">${esc(feat.Index)}</div>
          <div class="feat-title">${esc(feat.Text)}</div>
        </div>
        <div class="cfb-verdict">
          <div class="verdict-item"><div class="verdict-dot ${dc.dot}"></div><span class="${dc.verdict}">${esc(disclosure)}</span></div>
          <span class="verdict-sep">&middot;</span>
          <div class="verdict-item"><div class="verdict-dot ${fc.dot}"></div><span class="${fc.verdict}">${esc(essClass)}</span></div>
        </div>
        <div class="cfb-body">
          <div class="cfb-col">
            <h4>Analysis</h4>
            <div class="sub-heading">Interpretation</div>${renderAnalysisBlock(ana.Interpretation)}
            <div class="sub-heading">Mapping Summary</div>${renderAnalysisBlock(ana.Mapping_Summary)}
            <div class="sub-heading">Differences</div>${renderAnalysisBlock(ana.Differences)}
            <div class="sub-heading">Overall Opinion</div>${renderAnalysisBlock(ana.Overall_Opinion)}
            <div class="justification-panel">
              <div class="j-label">Essentiality Justification</div>
              <p>${esc(justification)}</p>
            </div>
          </div>
          <div class="cfb-col">
            <h4>Cited Standard Excerpts</h4>
            <div class="excerpts-section">
              <button class="excerpt-toggle" onclick="toggleExcerpts(this)" aria-expanded="false">
                <span class="toggle-left"><span>Standard Excerpts</span><span class="excerpt-count">${parsedExcs.length}</span></span>
                <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
              </button>
              <div class="excerpt-body">${excItemsHtml}</div>
            </div>
          </div>
        </div>
      </div>`;
    }).join("\n");
  }

  const LOGO_SVG = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.1" height="36" viewBox="0 0 3144.8497854077254 1027.5281652360513"><g transform="scale(7.242489270386266) translate(10, 10)"><defs id="SvgjsDefs1027"/><g id="SvgjsG1028" featureKey="symbolGroupContainer" transform="matrix(1.16515289568328,0,0,1.16515289568328,0.000007264315552150539,0.000007264315552150539)" fill="#fff"><path d="M52.3 104.6a52.3 52.3 0 1 1 52.3-52.3 52.4 52.4 0 0 1-52.3 52.3zm0-102.3a50 50 0 1 0 50 50 50 50 0 0 0-50-50z"/></g><g id="SvgjsG1029" featureKey="2ou6gm-0" transform="matrix(0.9971509971509972,0,0,0.9971509971509972,264.8062678062678,-335.4786324786325)" fill="#fff"><path d="M-167.5,390.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-165.5,389.6-166.4,390.5-167.5,390.5z M-177.5,428.5c-2.2,0-4-1.8-4-4s1.8-4,4-4c2.2,0,4,1.8,4,4S-175.3,428.5-177.5,428.5z M-177.5,410.5c-2.2,0-4-1.8-4-4s1.8-4,4-4c2.2,0,4,1.8,4,4S-175.3,410.5-177.5,410.5z M-177.5,392.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-173.5,390.7-175.3,392.5-177.5,392.5z M-177.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-173.5,372.7-175.3,374.5-177.5,374.5z M-194.5,414.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7c3.9,0,7,3.1,7,7C-187.5,411.4-190.6,414.5-194.5,414.5z M-194.5,394.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7c3.9,0,7,3.1,7,7C-187.5,391.4-190.6,394.5-194.5,394.5z M-195.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-191.5,372.7-193.3,374.5-195.5,374.5z M-195.5,362.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-193.5,361.6-194.4,362.5-195.5,362.5z M-214.5,414.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7s7,3.1,7,7C-207.5,411.4-210.6,414.5-214.5,414.5z M-214.5,394.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7s7,3.1,7,7C-207.5,391.4-210.6,394.5-214.5,394.5z M-213.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-209.5,372.7-211.3,374.5-213.5,374.5z M-213.5,362.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-211.5,361.6-212.4,362.5-213.5,362.5z M-231.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-227.5,372.7-229.3,374.5-231.5,374.5z M-231.5,384.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-235.5,386.3-233.7,384.5-231.5,384.5z M-241.5,408.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-239.5,407.6-240.4,408.5-241.5,408.5z M-241.5,390.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-239.5,389.6-240.4,390.5-241.5,390.5z M-231.5,402.5c2.2,0,4,1.8,4,4s-1.8,4-4,4c-2.2,0-4-1.8-4-4S-233.7,402.5-231.5,402.5z M-231.5,420.5c2.2,0,4,1.8,4,4s-1.8,4-4,4c-2.2,0-4-1.8-4-4S-233.7,420.5-231.5,420.5z M-213.5,420.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-217.5,422.3-215.7,420.5-213.5,420.5z M-213.5,432.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-215.5,433.4-214.6,432.5-213.5,432.5z M-195.5,420.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-199.5,422.3-197.7,420.5-195.5,420.5z M-195.5,432.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-197.5,433.4-196.6,432.5-195.5,432.5z M-167.5,404.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-169.5,405.4-168.6,404.5-167.5,404.5z" style="fill-rule:evenodd;clip-rule:evenodd;"/></g><g id="SvgjsG1030" featureKey="kZnDdN-0" transform="matrix(3.8775259911441498,0,0,3.8775259911441498,137.1154802767278,2.6123700442792526)" fill="#fff"><path d="M2.8906 8.457 c-0.88867 0 -1.6309 -0.72266 -1.6309 -1.6211 c0 -0.88867 0.74219 -1.6113 1.6309 -1.6113 c0.86914 0 1.6113 0.72266 1.6113 1.6113 c0 0.89844 -0.74219 1.6211 -1.6113 1.6211 z M1.4551 20 l0 -10.039 l2.832 0 l0 10.039 l-2.832 0 z M13.0859875 9.766 c2.6465 0 4.834 1.9434 4.834 5.2344 s-2.1875 5.2344 -4.834 5.2344 c-1.3086 0 -2.4805 -0.50781 -3.0762 -1.4258 l0 6.0742 l-2.8125 0 l0 -14.922 l2.666 0 l0.078125 1.3477 c0.55664 -0.99609 1.7773 -1.543 3.1445 -1.543 z M12.4511875 17.9004 c1.4746 0 2.6563 -1.0742 2.6563 -2.9004 s-1.1816 -2.9004 -2.6563 -2.9004 c-1.5039 0 -2.6758 1.1426 -2.6758 2.9004 s1.1719 2.9004 2.6758 2.9004 z M37.129296875 9.766 c2.1484 0 3.5352 1.0938 3.5352 3.1543 l0 7.0801 l-2.8125 0 l0 -6.2793 c0 -1.1816 -0.74219 -1.6992 -1.582 -1.6992 c-1.0059 0 -1.8945 0.57617 -1.8945 2.3145 l0 5.6641 l-2.8418 0 l0 -6.25 c0 -1.2012 -0.72266 -1.7285 -1.6113 -1.7285 c-0.97656 0 -1.8848 0.57617 -1.8848 2.4609 l0 5.5176 l-2.8027 0 l0 -10.039 l2.8027 0 l0 1.1816 c0.66406 -0.83008 1.7871 -1.3086 3.1152 -1.3086 z M44.833959375 8.457 c-0.88867 0 -1.6309 -0.72266 -1.6309 -1.6211 c0 -0.88867 0.74219 -1.6113 1.6309 -1.6113 c0.86914 0 1.6113 0.72266 1.6113 1.6113 c0 0.89844 -0.74219 1.6211 -1.6113 1.6211 z M43.398459375 20 l0 -10.039 l2.832 0 l0 10.039 l-2.832 0 z M55.068346875 9.766 c2.4121 0 3.7402 1.25 3.7402 3.4766 l0 6.7578 l-2.8223 0 l0 -6.1523 c0 -1.3379 -0.83008 -1.8262 -1.7969 -1.8262 c-1.1621 0 -2.2168 0.58594 -2.2363 2.4414 l0 5.5371 l-2.8125 0 l0 -10.039 l2.8125 0 l0 1.1133 c0.70313 -0.83008 1.7871 -1.3086 3.1152 -1.3086 z M68.652325 5 l2.8125 0 l0 15 l-2.666 0 l-0.068359 -1.3086 c-0.57617 0.98633 -1.7871 1.543 -3.1543 1.543 c-2.6465 0 -4.834 -1.9531 -4.834 -5.2344 s2.1973 -5.2344 4.834 -5.2344 c1.3184 0 2.4805 0.49805 3.0762 1.4063 l0 -6.1719 z M66.220725 17.9004 c1.4941 0 2.6563 -1.1426 2.6563 -2.9004 s-1.1719 -2.9102 -2.6563 -2.9102 c-1.4941 0 -2.666 1.1035 -2.666 2.9102 c0 1.7969 1.1719 2.9004 2.666 2.9004 z"/></g></g></svg>`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>${esc(patentNumber)} \u2013 Patent Analysis Report</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Source+Sans+3:wght@300;400;500;600&family=Source+Code+Pro:wght@400;500&display=swap" rel="stylesheet" />
  <style>
    :root{--brand:#ff6734;--brand-light:#fff0eb;--navy:#0f1f38;--ink:#1c1c2e;--mid:#4a4a6a;--muted:#7a7a96;--rule:#e2e2ed;--bg:#fafaf8;--surface:#ffffff;--surface-alt:#f4f4f0;--green:#1a6b4a;--green-bg:#eaf5ef;--amber:#8a5a00;--amber-bg:#fdf5e0;--red:#8a0000;--red-bg:#fdf0f0;--radius:6px;--radius-lg:12px;}
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
    html{scroll-behavior:smooth;}
    body{font-family:'Source Sans 3',sans-serif;font-size:15px;line-height:1.7;color:var(--ink);background:var(--bg);}
    @page{margin:2cm;}
    @media print{.excerpt-toggle{display:none;}.excerpt-body{display:block!important;}.methodology-toggle{display:none;}.methodology-body{display:block!important;}}
    .page-wrap{max-width:1040px;margin:0 auto;padding:0 32px 80px;}
    .brand-rule{height:4px;background:var(--brand);}
    .header-bar{background:var(--navy);}
    .header-bar-inner{max-width:1040px;margin:0 auto;padding:28px 32px;display:flex;justify-content:space-between;align-items:center;}
    .header-bar .logo svg{height:32px;width:auto;}
    .header-bar .confidential{font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.5);border:1px solid rgba(255,255,255,.2);padding:4px 12px;border-radius:2px;}
    .identity{padding:48px 0 40px;border-bottom:1px solid var(--rule);}
    .identity-meta{display:flex;gap:8px;align-items:center;margin-bottom:16px;}
    .pill{display:inline-block;font-size:11px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;padding:3px 10px;border-radius:2px;background:var(--brand-light);color:var(--brand);}
    .pill-navy{background:rgba(15,31,56,.08);color:var(--navy);}
    .pill-restricted{background:var(--amber-bg);color:var(--amber);}
    .restricted-notice{margin:24px 0 0;border-left:3px solid var(--amber);background:var(--amber-bg);padding:18px 24px;border-radius:0 var(--radius) var(--radius) 0;}
    .restricted-label{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--amber);margin-bottom:10px;}
    .restricted-notice p{font-size:13px;line-height:1.75;color:var(--amberText,#8a5a00);}
    .identity h1{font-family:'Playfair Display',serif;font-size:32px;font-weight:600;color:var(--navy);line-height:1.2;margin-bottom:8px;}
    .identity-grid{display:grid;grid-template-columns:repeat(3,1fr);border:1px solid var(--rule);border-radius:var(--radius);overflow:hidden;margin-top:32px;}
    .identity-cell{padding:16px 20px;border-right:1px solid var(--rule);}
    .identity-cell:last-child{border-right:none;}
    .identity-cell .label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:4px;}
    .identity-cell .value{font-size:14px;font-weight:600;color:var(--navy);}
    .claim-block{margin:36px 0 0;border-left:3px solid var(--brand);background:var(--surface);padding:20px 24px;border-radius:0 var(--radius) var(--radius) 0;}
    .claim-block .claim-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--brand);margin-bottom:10px;}
    .claim-block p{font-size:14px;line-height:1.75;color:var(--ink);font-style:italic;}
    .section{margin-top:56px;}
    .section-heading{display:flex;align-items:center;gap:16px;margin-bottom:28px;}
    .section-heading h2{font-family:'Playfair Display',serif;font-size:22px;font-weight:600;color:var(--navy);white-space:nowrap;}
    .section-heading::after{content:'';flex:1;height:1px;background:var(--rule);}
    .summary-cards{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:16px;}
    .summary-cards.two-col{grid-template-columns:1fr 3fr;}
    .summary-card{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius);padding:18px 20px;}
    .summary-card.highlight{background:var(--amber-bg);border-color:#e8c96a;}
    .summary-card.highlight-green{background:var(--green-bg);border-color:#a0d4b8;}
    .summary-card .sc-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:8px;}
    .summary-card .sc-value{font-size:22px;font-weight:700;color:var(--navy);line-height:1.1;}
    .summary-card .sc-value.amber{font-size:15px;color:var(--amber);}
    .summary-card .sc-value.green{font-size:15px;color:var(--green);}
    .summary-card .sc-value.red{font-size:15px;color:var(--red);}
    .summary-card .sc-value.meta{font-size:14px;font-weight:500;color:var(--mid);}
    .opinion-box{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius-lg);padding:28px 32px;}
    .opinion-box h3{font-family:'Playfair Display',serif;font-size:16px;font-weight:600;color:var(--navy);margin-bottom:12px;}
    .opinion-box p{font-size:14px;line-height:1.8;color:var(--mid);}
    .limitations-box{background:var(--surface-alt);border:1px solid var(--rule);border-radius:var(--radius);padding:20px 24px;margin-top:16px;}
    .limitations-box .lim-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:8px;}
    .limitations-box p{font-size:14px;line-height:1.8;color:var(--mid);}
    .mapping-list{display:flex;flex-direction:column;gap:16px;margin-top:8px;}
    .mapping-item{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius);overflow:hidden;}
    .mapping-item-header{display:flex;align-items:flex-start;gap:16px;padding:16px 20px;}
    .feat-num{flex-shrink:0;width:26px;height:26px;border-radius:50%;background:var(--navy);color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:2px;}
    .mapping-item-header .feat-text{flex:1;font-size:14px;font-weight:600;color:var(--ink);line-height:1.5;}
    .badge{flex-shrink:0;display:inline-block;font-size:11px;font-weight:600;padding:3px 10px;border-radius:2px;}
    .badge-amber{background:var(--amber-bg);color:var(--amber);}
    .badge-green{background:var(--green-bg);color:var(--green);}
    .badge-red{background:var(--red-bg);color:var(--red);}
    .mapping-item-body{border-top:1px solid var(--rule);padding:14px 20px 14px 62px;background:var(--surface-alt);}
    .mapping-item-body p{font-size:13.5px;line-height:1.75;color:var(--mid);margin-bottom:8px;}
    .mapping-item-body p:last-child{margin-bottom:0;}
    .claim-feature-block{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius-lg);overflow:hidden;margin-bottom:32px;}
    .cfb-header{background:var(--navy);padding:20px 28px;display:flex;align-items:flex-start;gap:16px;}
    .cfb-header .feat-num{background:var(--brand);font-size:13px;width:28px;height:28px;flex-shrink:0;margin-top:1px;}
    .cfb-header .feat-title{font-size:14px;font-weight:500;color:rgba(255,255,255,.9);line-height:1.55;flex:1;font-style:italic;}
    .cfb-verdict{display:flex;gap:10px;padding:16px 28px;background:var(--surface-alt);border-bottom:1px solid var(--rule);}
    .verdict-item{display:flex;align-items:center;gap:8px;}
    .verdict-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
    .dot-amber{background:#d4a00a;}.dot-green{background:#1a6b4a;}.dot-red{background:#8a0000;}
    .verdict-item span{font-size:12px;font-weight:600;letter-spacing:.05em;text-transform:uppercase;}
    .verdict-item span.amber{color:var(--amber);}.verdict-item span.green{color:var(--green);}.verdict-item span.red{color:var(--red);}
    .verdict-sep{color:var(--rule);margin:0 4px;}
    .cfb-body{display:grid;grid-template-columns:1fr 1fr;}
    .cfb-col{padding:24px 28px;border-right:1px solid var(--rule);}
    .cfb-col:last-child{border-right:none;}
    .cfb-col h4{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--rule);}
    .cfb-col p{font-size:13.5px;line-height:1.75;color:var(--mid);margin-bottom:10px;}
    .cfb-col p:last-child{margin-bottom:0;}
    .sub-heading{font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--navy);margin:18px 0 8px;}
    .sub-heading:first-of-type{margin-top:0;}
    .justification-panel{margin:20px 0 0;background:var(--surface);border:1px solid var(--rule);border-left:3px solid var(--brand);border-radius:0 var(--radius) var(--radius) 0;padding:18px 22px;}
    .justification-panel .j-label{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--brand);margin-bottom:8px;}
    .justification-panel p{font-size:13.5px;line-height:1.75;color:var(--mid);}
    .excerpt-toggle{width:100%;background:none;border:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;padding:14px 0;color:var(--navy);font-family:'Source Sans 3',sans-serif;font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;transition:color .15s;}
    .excerpt-toggle:hover{color:var(--brand);}
    .excerpt-toggle .toggle-left{display:flex;align-items:center;gap:10px;}
    .excerpt-count{background:var(--navy);color:#fff;font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px;}
    .chevron{width:16px;height:16px;transition:transform .2s;color:var(--muted);}
    .chevron.open{transform:rotate(180deg);}
    .excerpt-body{display:none;}.excerpt-body.open{display:block;}
    .excerpt-item{margin-top:16px;border:1px solid var(--rule);border-radius:var(--radius);overflow:hidden;}
    .excerpt-item-header{background:var(--surface-alt);padding:10px 16px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--rule);}
    .exc-num{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--navy);}
    .exc-ref{font-size:11px;color:var(--muted);font-family:'Source Code Pro',monospace;}
    .excerpt-item-body{padding:14px 16px;background:#fafafa;font-family:'Source Code Pro',monospace;font-size:12px;line-height:1.65;color:var(--mid);overflow-x:auto;}
    .excerpt-item-body h4{font-family:'Source Sans 3',sans-serif;font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--navy);margin-bottom:10px;}
    .excerpt-item-body p{margin-bottom:6px;}.excerpt-item-body p:last-child{margin-bottom:0;}
    .exc-subhead{font-weight:600;color:var(--navy);margin-top:10px!important;}
    .exc-indent{padding-left:20px;}
    .exc-list{padding-left:18px;margin:6px 0;}.exc-list li{margin-bottom:4px;}
    .exc-table{width:100%;border-collapse:collapse;font-size:11px;margin:8px 0;}
    .exc-table th,.exc-table td{border:1px solid var(--rule);padding:5px 8px;text-align:left;vertical-align:top;}
    .exc-table th{background:var(--surface-alt);font-weight:600;color:var(--navy);}
    /* ── Methodology Key to Terms ──────────────────────────────────────── */
    .methodology-panel{border:1px solid var(--rule);border-radius:var(--radius);overflow:hidden;margin-top:16px;}
    .methodology-toggle{width:100%;background:var(--surface-alt);border:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;padding:13px 20px;color:var(--navy);font-family:'Source Sans 3',sans-serif;font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;transition:background .15s;}
    .methodology-toggle:hover{background:#ebebea;}
    .meth-toggle-left{display:flex;align-items:center;gap:8px;}
    .methodology-body{display:none;border-top:1px solid var(--rule);}.methodology-body.open{display:block;}
    .meth-group-header{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);padding:12px 20px 6px;background:var(--surface);border-bottom:1px solid var(--rule);}
    .meth-row{display:grid;grid-template-columns:200px 1fr;border-bottom:1px solid var(--rule);}
    .meth-row:last-child{border-bottom:none;}
    .meth-label{padding:10px 14px;font-size:12px;font-weight:700;line-height:1.45;border-right:1px solid var(--rule);display:flex;align-items:flex-start;}
    .meth-label-green{color:var(--green);background:var(--green-bg);}
    .meth-label-amber{color:var(--amber);background:var(--amber-bg);}
    .meth-label-red{color:var(--red);background:var(--red-bg);}
    .meth-label-navy{color:var(--navy);background:var(--surface-alt);}
    .meth-def{padding:10px 16px;font-size:13px;line-height:1.65;color:var(--mid);background:var(--surface);}
    /* ── Disclaimer ────────────────────────────────────────────────────── */
    .disclaimer{margin-top:64px;border:1px solid var(--rule);border-radius:var(--radius-lg);overflow:hidden;}
    .disclaimer-header{background:var(--surface-alt);padding:16px 24px;border-bottom:1px solid var(--rule);display:flex;align-items:center;gap:10px;}
    .disclaimer-icon{width:16px;height:16px;color:var(--muted);}
    .disclaimer-header h4{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);}
    .disclaimer-body{padding:20px 24px;}
    .disclaimer-body ol{padding-left:18px;display:flex;flex-direction:column;gap:10px;}
    .disclaimer-body li{font-size:12.5px;line-height:1.7;color:var(--muted);}
    .disclaimer-body li strong{font-weight:600;color:var(--mid);}
    .site-footer{text-align:center;padding:32px 0 0;font-size:12px;color:var(--muted);letter-spacing:.08em;}
    @media(max-width:760px){.page-wrap{padding:0 16px 60px;}.summary-cards{grid-template-columns:1fr 1fr;}.summary-cards.two-col{grid-template-columns:1fr;}.identity-grid{grid-template-columns:1fr;}.identity-cell{border-right:none;border-bottom:1px solid var(--rule);}.cfb-body{grid-template-columns:1fr;}.cfb-col{border-right:none;border-bottom:1px solid var(--rule);}.meth-row{grid-template-columns:1fr;}.meth-label{border-right:none;border-bottom:1px solid var(--rule);}}
  </style>
</head>
<body>
  <div class="brand-rule"></div>
  <div class="header-bar">
    <div class="header-bar-inner">
      <div class="logo">${LOGO_SVG}</div>
      <div class="confidential">Confidential</div>
    </div>
  </div>
  <div class="page-wrap">
    <div class="identity">
      <div class="identity-meta">
        <span class="pill">${esc(patentNumber)}</span>
        <span class="pill pill-navy">${esc(standard)}</span>
        <span class="pill pill-navy">${esc(claimNumber)} &middot; ${esc(claimCategory)}</span>
        ${restricted ? '<span class="pill pill-restricted">Restricted Use</span>' : ""}
      </div>
      <h1>${esc(title)}</h1>
      <div class="identity-grid">
        <div class="identity-cell"><div class="label">Patent Number</div><div class="value">${esc(patentNumber)}</div></div>
        <div class="identity-cell"><div class="label">Owner</div><div class="value">${esc(owner)}</div></div>
        <div class="identity-cell"><div class="label">Standard</div><div class="value">${esc(standard)}</div></div>
      </div>
      <div class="claim-block">
        <div class="claim-label">${esc(claimLabel)}</div>
        <p>${esc(claimText)}</p>
      </div>
      ${restricted ? `
      <div class="restricted-notice">
        <div class="restricted-label">&#9888;&nbsp; Restricted Use Notice</div>
        <p>${RESTRICTED_NOTICE_HTML}</p>
      </div>` : ""}
    </div>
    <div class="section">
      <div class="section-heading"><h2>Executive Summary</h2></div>
      <div class="summary-cards">
        <div class="summary-card"><div class="sc-label">Claim Number</div><div class="sc-value">${esc(claimNumber)}</div></div>
        <div class="summary-card"><div class="sc-label">Claim Category</div><div class="sc-value">${esc(claimCategory)}</div></div>
        <div class="summary-card"><div class="sc-label">Percentage Mapped</div><div class="sc-value">${esc(pctMapped)}</div></div>
        <div class="summary-card ${ec.card}"><div class="sc-label">Essentiality Decision</div><div class="sc-value ${ec.value}">${esc(essDecision)}</div></div>
      </div>
      <div class="summary-cards two-col">
        <div class="summary-card"><div class="sc-label">Weighted Mapping</div><div class="sc-value">${esc(pctWeighted)}</div></div>
        <div class="summary-card"><div class="sc-label">Limitations</div><div class="sc-value meta">${esc(limLabel)}</div></div>
      </div>
      ${buildMethodologyHtml(methodology)}
      <div class="opinion-box" style="margin-top:16px;">
        <h3>Opinion</h3>
        <p>${esc(opinion)}</p>
        <div class="limitations-box">
          <div class="lim-label">Limitations Detail</div>
          <p>${esc(limBody)}</p>
        </div>
      </div>
      <div style="margin-top:32px;">
        <div class="section-heading" style="margin-top:0;"><h2>Mapping Summary</h2></div>
        <div class="mapping-list">${buildMappingSummaryHtml(mappingItems)}</div>
      </div>
    </div>
    <div class="section">
      <div class="section-heading"><h2>Claim Chart</h2></div>
      ${buildClaimChartHtml(charts)}
    </div>
    <div class="disclaimer">
      <div class="disclaimer-header">
        <svg class="disclaimer-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        <h4>Disclaimer</h4>
      </div>
      <div class="disclaimer-body">
        <ol>
          <li><strong>Preliminary and Informational Nature:</strong> The present work product was generated using a prototype AI model and is provided for informational purposes only. It does not constitute a legal or technical opinion regarding the essentiality or non-essentiality of any patent claim to any technical standard. It is not a substitute for legal or technical advice, and clients are strongly encouraged to seek independent professional counsel before relying on this material for purposes such as licensing, enforcement, or infringement analysis.</li>
          <li><strong>Scope of Analysis:</strong> The analysis is limited to the individual patent claim(s) identified in the chart and does not take into account the full patent specification, including the description and drawings. Consequently, any interpretation of claim scope is based on the claim language alone and may differ from that reached through a full legal construction under applicable law.</li>
          <li><strong>Referencing of Standards:</strong> Where citations to section numbers, table numbers, or figure numbers in a technical standard are provided, they are included for convenience only. While care is taken in referencing, these citations should not be relied upon as authoritative without verification against the official version of the standard.</li>
          <li><strong>Interpretation of Standards:</strong> References to technical standards are based on publicly available documents. Where relevant, excerpts are cited in text form. Figures and diagrams from such standards are not reproduced; instead, any associated visual content is paraphrased using descriptive language. Such paraphrasing should not be construed as a verbatim or authoritative interpretation of the standard itself.</li>
          <li><strong>Subjectivity of Essentiality:</strong> Determinations of potential alignment between a patent claim and a standard may depend on how specific terms or functional steps are construed. What may appear to correspond closely under one interpretation may be viewed as merely analogous under another. This assessment is inherently interpretive and does not reflect a consensus view or judicial determination.</li>
          <li><strong>Implementation Considerations:</strong> The presence of a feature in a standard does not imply that all compliant implementations necessarily use that feature. A compliant product may omit or bypass specific technical elements referenced in a patent claim.</li>
          <li><strong>Alternative Solutions:</strong> Standards may include multiple options or alternative techniques to achieve similar functionality. A given patent claim may correspond to one such option, but not to others that are also compliant with the standard.</li>
          <li><strong>Legal Proceedings:</strong> In the context of litigation, essentiality determinations typically require a far more detailed analysis, including expert testimony, claim construction under applicable law, and examination of implementation evidence. The present assessment should not be relied upon for litigation, licensing negotiation, or investment decisions without further professional review.</li>
        </ol>
      </div>
    </div>
    <div class="site-footer">ipmind.ai</div>
  </div>
  <script>
    function toggleExcerpts(btn) {
      const body = btn.nextElementSibling;
      const chevron = btn.querySelector('.chevron');
      const isOpen = body.classList.contains('open');
      body.classList.toggle('open', !isOpen);
      chevron.classList.toggle('open', !isOpen);
      btn.setAttribute('aria-expanded', String(!isOpen));
    }
    function toggleMethodology(btn) {
      const body = btn.nextElementSibling;
      const chevron = btn.querySelector('.chevron');
      const isOpen = body.classList.contains('open');
      body.classList.toggle('open', !isOpen);
      chevron.classList.toggle('open', !isOpen);
      btn.setAttribute('aria-expanded', String(!isOpen));
    }
  </script>
</body>
</html>`;
}

// ═════════════════════════════════════════════════════════════════════════
// EXPRESS ROUTES — /generate-html
// ═════════════════════════════════════════════════════════════════════════

app.post("/generate-html", (req, res) => {
  try {
    const body = req.body;
    const data = Array.isArray(body) ? body[0] : body;
    const meta = {
      Patent_Number: req.query.patent   || "",
      Title:         req.query.title    || "",
      Owner:         req.query.owner    || "",
      Standard:      req.query.standard || "",
    };
    const html = buildHtml(data, meta, req.query.restricted === "true" || data.Restricted_Use === true);
    const safeName = (data.Patent_Number || meta.Patent_Number || "report")
      .replace(/[^A-Za-z0-9_-]/g, "_");
    res.json({
      html,
      filename: safeName + "_report.html",
      patent:   data.Patent_Number || meta.Patent_Number || "",
      claim:    data.Claim_Number  || "",
      features: (data.Claim_Charts || []).length,
      excerpts_total: (data.Claim_Charts || []).reduce((s, c) => s + (c.Cited_Excerpts || []).length, 0),
    });
  } catch (err) {
    console.error("Error generating html:", err);
    res.status(500).json({ error: err.message });
  }
});

// ═════════════════════════════════════════════════════════════════════════
// START
// ═════════════════════════════════════════════════════════════════════════

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("IPMIND docx service running on port " + PORT));
