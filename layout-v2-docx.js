"use strict";
// ─────────────────────────────────────────────────────────────────────────────
// Claim chart layout v2 (Word) — ipmind-docx-service
// Word counterpart of layout-v2.js. Rendered when /generate is called with
// ?layout=v2 and the payload carries a Snapshot; otherwise layout v1.
// Sections: 1 (portrait) identity, at a glance, summary mapping;
//           2 (landscape) detailed claim chart with element colour highlights;
//           3 (portrait) appendix: conclusion, limitations, key to terms,
//             restricted use notice, disclaimer.
// Helpers (header, footer, disclaimer, notice, palette) are injected from
// server.js so both layouts share them.
// ─────────────────────────────────────────────────────────────────────────────
const { locate, parseExcerptV2 } = require("./layout-v2");

module.exports = function makeBuildDocumentV2(h) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
          ShadingType, SectionType, PageOrientation, VerticalAlign } = h.docx;
  const { C, PG, PGL, safeStr, shade, solidBorder, noBorders, noBorder, emptyPara,
          sectionHeading, makeHeader, makeFooter, disclaimerSection, restrictedNoticePage } = h;

  const EL   = ["C9DAF6", "E2D1F7", "C0E8E0", "F6D1E2", "D9E9C4"];
  const ELN  = ["2A58A3", "6A3FA8", "1D6F67", "A03468", "4E7A1E"];
  const BADGE = {
    green: ["EAF5EF", "1A6B4A"], core: ["1A6B4A", "FFFFFF"], amber: ["FDF5E0", "8A5A00"],
    red: ["FDF0F0", "8A0000"], grey: ["ECECF2", "4A4A6A"], navy: ["0F1F38", "FFFFFF"],
  };
  const arr = (x) => Array.isArray(x) ? x : (x == null ? [] : [x]);
  const T = (text, o = {}) => new TextRun({ text: safeStr(text), font: "Arial", size: 19, color: C.ink, ...o });
  const P = (children, o = {}) => new Paragraph({ children: arr(children), spacing: { after: 80 }, ...o });
  const badge = (text, kind) => {
    const [fill, color] = BADGE[kind] || BADGE.grey;
    return [T(" " + text + " ", { size: 16, bold: true, color, shading: { type: ShadingType.CLEAR, fill, color: "auto" } }), T("  ", { size: 16 })];
  };
  const cell = (children, w, o = {}) => new TableCell({
    children: arr(children).length ? arr(children) : [emptyPara()], width: { size: w, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    borders: { top: solidBorder(C.rule, 4), bottom: solidBorder(C.rule, 4), left: noBorder, right: noBorder }, ...o,
  });
  const table = (W, widths, rows) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows });
  const headRow = (labels, widths) => new TableRow({ tableHeader: true, children: labels.map((l, k) =>
    cell([P(T(l, { size: 16, bold: true, color: C.muted }), { spacing: { after: 0 } })], widths[k], { shading: shade(C.surfaceAlt) })) });
  const sub = (text) => new Paragraph({ children: [T(text, { size: 22, bold: true, color: C.navy })], spacing: { before: 280, after: 100 } });

  const shortDisc = (d) => ({ "Explicitly Disclosed": "Explicit", "Implicitly Disclosed": "Implicit", "Functionally Equivalent": "Equivalent", "Not Disclosed": "Not disclosed" }[d] || d || "");
  const discKind  = (d) => d === "Not Disclosed" ? "red" : d === "Functionally Equivalent" ? "amber" : "green";
  function essParts(cls) {
    const [scope, side] = String(cls || "").split(" | ");
    const m = scope.match(/^Essential \((.+)\)$/);
    const label = m ? m[1].replace(/ in Main Profile$/, "").replace(/^Optional in Profile SEP$/, "Optional (profile SEP)") : scope;
    const kind = m ? (/^Core/.test(m[1]) ? "core" : "green") : (/inherent|non-technical/i.test(scope) ? "grey" : /implementation/i.test(scope) ? "amber" : "red");
    return { scope, side: side || "", label, kind };
  }
  const sideWords = (s) => ({ "Encoder and Decoder": "Encoder and decoder", "Decoder": "Decoder only", "Encoder": "Encoder only", "Codec System": "Codec system only" }[s] || "");
  const featList = (xs) => { const a = arr(xs).map(String); return !a.length ? "" : a.length === 1 ? `feature ${a[0]}` : `features ${a.slice(0, -1).join(", ")} and ${a[a.length - 1]}`; };
  const cap = (s) => s ? s[0].toUpperCase() + s.slice(1) : s;

  // Turns text + ranges into paragraphs of runs; highlighted runs are shaded
  // in the element colour and followed by a superscript element number.
  function highlightedParas(text, ranges, base, paraOpts = {}) {
    const t = String(text ?? "");
    const rs = ranges.filter(Boolean).sort((a, b) => a.s - b.s);
    const segs = []; let pos = 0;
    for (const r of rs) { if (r.s < pos) continue; if (r.s > pos) segs.push({ t: t.slice(pos, r.s) }); segs.push({ t: t.slice(r.s, r.e), k: r.k, n: r.n }); pos = r.e; }
    if (pos < t.length) segs.push({ t: t.slice(pos) });
    const paras = []; let runs = [];
    const flush = () => { paras.push(new Paragraph({ children: runs.length ? runs : [T("", base)], spacing: { after: 0 }, ...paraOpts })); runs = []; };
    for (const sg of segs) {
      const lines = sg.t.split("\n");
      lines.forEach((ln, li) => {
        if (li > 0) flush();
        if (ln) runs.push(T(ln, sg.k === undefined ? base : { ...base, shading: { type: ShadingType.CLEAR, fill: EL[sg.k], color: "auto" } }));
      });
      if (sg.k !== undefined) runs.push(T(String(sg.n), { ...base, superScript: true, bold: true, color: ELN[sg.k] }));
    }
    flush();
    return paras;
  }

  return async function buildDocumentV2(data, meta, restricted) {
    const snap = data.Snapshot || {}, verdict = snap.verdict || {}, align = snap.alignment || {};
    const charts = arr(data.Claim_Charts);
    const sumByIdx = {}; for (const f of arr(snap.features)) sumByIdx[String(f.index)] = f;
    const patent = safeStr(data.Patent_Number || meta.Patent_Number || "");
    const title  = safeStr(data.Title || meta.Title || "Patent analysis report");
    const owner  = safeStr(data.Owner || meta.Owner || "");
    const standard = [data.Standard || meta.Standard || "", data.Standard_Edition || ""].filter(Boolean).join(" · ");
    const claimNo = safeStr(data.Claim_Number || "");
    const catWords = { Encoder: "Encoder claim", Decoder: "Decoder claim", Encoder_and_Decoder_System: "Encoder and decoder claim", Other: "Other claim type" }[data.Claim_Category] || "";

    // ── Section 1: identity, at a glance, summary mapping ──
    const s1 = [];
    s1.push(new Paragraph({ children: [
      ...(patent ? badge(patent, "grey") : []), ...(claimNo ? badge("Claim " + claimNo, "grey") : []),
      ...(standard ? badge(standard, "navy") : []), ...(restricted ? badge("RESTRICTED USE", "amber") : []) ], spacing: { before: 240, after: 120 } }));
    s1.push(new Paragraph({ children: [new TextRun({ text: title, font: "Georgia", size: 48, bold: true, color: C.navy })], spacing: { after: 80 } }));
    if (owner || catWords) s1.push(P(T([owner, catWords].filter(Boolean).join(" · "), { color: C.muted })));
    if (data.Claim) s1.push(h.claimBlock("Claim " + claimNo, data.Claim));

    s1.push(...sectionHeading("At a glance"));
    const ov = essParts((verdict.classification || "") + (verdict.side ? " | " + verdict.side : ""));
    const essential = ov.kind === "core" || ov.kind === "green";
    const vFill = essential ? "EAF5EF" : ov.kind === "amber" ? "FDF5E0" : ov.kind === "grey" ? C.surfaceAlt : "FDF0F0";
    const vText = essential ? "1A6B4A" : ov.kind === "amber" ? "8A5A00" : ov.kind === "grey" ? C.navy : "8A0000";
    const vLabel = essential ? "Essential · " + ov.scope.replace(/^Essential \((.+)\)$/, "$1") : (ov.scope || "—");
    const third = Math.floor(PG.W / 3);
    s1.push(table(PG.W, [third, third, PG.W - 2 * third], [
      new TableRow({ children: [cell([
        new Paragraph({ children: [new TextRun({ text: vLabel, font: "Georgia", size: 30, bold: true, color: vText })], spacing: { after: 80 } }),
        ...(verdict.bottom_line ? [P(T(verdict.bottom_line, { size: 20 }), { spacing: { after: 0 } })] : []) ], PG.W, { columnSpan: 3, shading: shade(vFill), borders: noBorders })] }),
      new TableRow({ children: [
        [`${verdict.percentage_mapped ?? "—"}%`, "Mapped"], [`${verdict.weighted_percentage_mapped ?? "—"}%`, "Evidence strength"],
        [`${verdict.features_aligned ?? "—"}/${verdict.features_technical ?? "—"}`, "Features aligned"] ].map(([v, l], k) =>
        cell([P(T(v, { size: 32, bold: true, color: C.navy }), { spacing: { after: 0 } }), P(T(l, { size: 16, color: C.mid }), { spacing: { after: 0 } })],
             k < 2 ? third : PG.W - 2 * third, { shading: shade(vFill), borders: noBorders })) }),
    ]));

    // Where it aligns
    const parts = [];
    if (arr(align.stated_directly).length) parts.push(`${cap(featList(align.stated_directly))} ${align.stated_directly.length > 1 ? "are" : "is"} stated directly in the standard`);
    if (arr(align.by_implication).length)  parts.push(`${featList(align.by_implication)} ${align.by_implication.length > 1 ? "follow" : "follows"} by necessary implication`);
    if (arr(align.by_equivalence).length)  parts.push(`${featList(align.by_equivalence)} ${align.by_equivalence.length > 1 ? "map" : "maps"} by functional equivalence`);
    if (arr(align.not_disclosed).length)   parts.push(`${featList(align.not_disclosed)} ${align.not_disclosed.length > 1 ? "are" : "is"} not disclosed`);
    if (arr(align.inherent).length)        parts.push(`${featList(align.inherent)} ${align.inherent.length > 1 ? "are inherent components" : "is an inherent component"}`);
    s1.push(sub("Where it aligns"));
    if (parts.length) s1.push(P(T(cap(parts.join("; ")) + ".", { color: C.mid })));
    const pw = [1500, PG.W - 1500 - 1800, 1800];
    const provs = arr(align.provisions);
    if (provs.length) s1.push(table(PG.W, pw, [headRow(["Clause", "What it provides", "Features"], pw), ...provs.map(p => new TableRow({ children: [
      cell(P(T(p.clause, { bold: true, font: "Courier New", size: 18 }), { spacing: { after: 0 } }), pw[0]),
      cell(P([T(p.title), ...(p.verify ? [T("  "), ...badge("Verify", "amber")] : [])], { spacing: { after: 0 } }), pw[1]),
      cell(P(T(arr(p.features).map(f => "Feature " + f).join(", "), { size: 17 }), { spacing: { after: 0 } }), pw[2]) ] })) ]));

    // Gaps and conditions
    s1.push(sub("Gaps and conditions"));
    const G = arr(snap.gaps_and_conditions);
    const gk = { Gap: "red", Equivalence: "amber", Construction: "amber", Condition: "green", Verify: "amber", "Non-required": "red", Product: "amber" };
    const gw = [1700, PG.W - 1700 - 1500, 1500];
    const gRows = [];
    if (!G.some(g => g.type === "Gap")) gRows.push(["Gap", "None", ""]);
    for (const g of G) gRows.push([g.type, g.type === "Condition" && arr(g.flags).length ? (arr(g.tools).length ? `${g.text} (${arr(g.flags).join(" / ")})` : "Only when this enabling flag is set: " + arr(g.flags).join(" / ")) : (g.type === "Construction" && g.claim_words ? `${g.question || ("How should “" + g.claim_words + "” be read?")} Broad reading: ${g.adopted_result}. Narrower: ${g.narrower_result}.` : g.text), g.feature ? "Feature " + g.feature : ""]);
    s1.push(table(PG.W, gw, gRows.map(([ty, tx, f]) => new TableRow({ children: [
      cell(P(badge(ty, gk[ty] || "amber"), { spacing: { after: 0 } }), gw[0]), cell(P(T(tx), { spacing: { after: 0 } }), gw[1]),
      cell(P(T(f, { size: 17 }), { spacing: { after: 0 } }), gw[2]) ] }))));

    // Summary mapping
    s1.push(...sectionHeading("Summary mapping"));
    s1.push(P(T("One line per feature. The full analysis and evidence for each is in the detailed claim chart.", { color: C.mid, size: 18 })));
    const sw = [450, 2500, 2600, 1900, PG.W - 450 - 2500 - 2600 - 1900];
    const caveatOf = (gap, disc) => gap && gap !== "None" ? gap : disc === "Not Disclosed" ? "Not disclosed in the standard" : disc === "Functionally Equivalent" ? "Relies on functional equivalence" : "";
    s1.push(table(PG.W, sw, [headRow(["#", "Claim feature", "Maps to", "Essentiality", "Gap or caveat"], sw), ...charts.map(c => {
      const i = String(c?.Claim_Feature?.Index ?? ""), s = sumByIdx[i] || {}, dec = c.Decision || {}, e = essParts(dec.Essentiality_Classification);
      const flags = arr(dec.Gating_Flags);
      const cpS = (dec.Construction && dec.Construction.claim_words) ? dec.Construction : null;
      const tools = [...new Set(arr(dec.Condition_Tools).map(x => x && x.tool).filter(Boolean))];
      return new TableRow({ cantSplit: true, children: [
        cell(P(T(i, { bold: true }), { spacing: { after: 0 } }), sw[0]),
        cell(P(T(s.short_label || c?.Claim_Feature?.Text || ""), { spacing: { after: 0 } }), sw[1]),
        cell(P(T(s.maps_to || "", { size: 17, color: C.mid }), { spacing: { after: 0 } }), sw[2]),
        cell([P(badge(e.label, e.kind), { spacing: { after: 0 } }), ...(flags.length ? [P([...(tools.length ? [T("Needs " + tools.join(" and ") + " ", { size: 15, color: C.mid })] : []), T(flags.join(" / "), { size: 15, font: "Courier New", color: C.mid })], { spacing: { after: 0 } })] : []),
              ...(cpS ? [P(T(`Narrower reading: ${cpS.narrower_result}`, { size: 15, color: "8A5A00" }), { spacing: { after: 0 } })] : [])], sw[3]),
        cell([...(caveatOf(s.gap, dec.Disclosure) || !cpS ? [P(T(caveatOf(s.gap, dec.Disclosure) || "None", { size: 17, color: caveatOf(s.gap, dec.Disclosure) ? "8A5A00" : C.muted }), { spacing: { after: 0 } })] : []),
              ...(cpS ? [P([...badge("Construction", "amber"), T(cpS.question || `“${cpS.claim_words}” could be read more narrowly`, { size: 16 })], { spacing: { after: 0 } })] : [])], sw[4]) ] });
    })]));

    // ── Section 2: detailed claim chart (landscape) ──
    const s2 = [...sectionHeading("Detailed claim chart"),
      P(T("Each element of a feature has its own colour and number. The same colour and number mark the supporting text in the excerpts.", { color: C.mid, size: 18 }))];
    charts.forEach((c, ci) => {
      const i = String(c?.Claim_Feature?.Index ?? ""), s = sumByIdx[i] || {}, dec = c.Decision || {}, an = c.Analysis || {};
      const e = essParts(dec.Essentiality_Classification);
      const ftext = c?.Claim_Feature?.Text || "";
      const ex = arr(c.Cited_Excerpts).map(parseExcerptV2); const exByNum = {}; ex.forEach(x => exByNum[x.num] = x);
      let k = 0;
      const colored = arr(dec.Elements).map(el => el.status === "Framing" ? { el, n: 0 } : { el, n: ++k, k: (k - 1) % 5 });
      const ftR = colored.filter(x => x.n).map(x => { const r = locate(ftext, x.el.element_text) || locate(ftext, String(x.el.element_text || "").replace(/^\s*\d+\.\s*/, "")); return r ? { s: r[0], e: r[1], k: x.k, n: x.n } : null; });
      const exR = {}, cited = new Set();
      for (const x of colored) for (const ev of arr(x.el.evidence)) {
        const num = String(ev.excerpt_number); cited.add(num);
        const tx = exByNum[num]; if (!tx || !x.n) continue;
        const r = locate(tx.text, ev.quote); if (r) (exR[num] = exR[num] || []).push({ s: r[0], e: r[1], k: x.k, n: x.n });
      }
      if (ci > 0) s2.push(emptyPara());
      // feature header
      s2.push(table(PGL.W, [700, PGL.W - 700], [new TableRow({ children: [
        cell(P(T(i, { bold: true, color: "FFFFFF", size: 24 }), { spacing: { after: 0 } }), 700, { shading: shade(C.orange), borders: noBorders, verticalAlign: VerticalAlign.CENTER }),
        cell([...highlightedParas(ftext, ftR, { size: 22 }), P([...badge(dec.Disclosure || "", discKind(dec.Disclosure)), ...badge(e.label, e.kind),
          ...(arr(dec.Gating_Flags).length ? [T(arr(dec.Gating_Flags).join(" / ") + "   ", { size: 15, font: "Courier New", color: C.mid })] : []),
          T(sideWords(e.side), { size: 16, color: C.mid })], { spacing: { before: 80, after: 0 } })], PGL.W - 700,
          { borders: { top: solidBorder(C.rule, 4), bottom: solidBorder(C.rule, 4), left: noBorder, right: solidBorder(C.rule, 4) } }) ] })]));
      // element table
      if (colored.length) {
        const ew = [600, 3600, PGL.W - 600 - 3600 - 2000 - 1600, 2000, 1600];
        s2.push(new Paragraph({ children: [T("How each element maps", { bold: true, color: C.navy })], spacing: { before: 160, after: 80 } }));
        s2.push(table(PGL.W, ew, [headRow(["", "Claim element", "Maps to", "Disclosure", "Evidence"], ew), ...colored.map(x => new TableRow({ cantSplit: true, children: [
          cell(P(T(x.n ? String(x.n) : "", { bold: true, color: x.n ? ELN[x.k] : C.muted }), { spacing: { after: 0 } }), ew[0], { shading: shade(x.n ? EL[x.k] : "ECECF2") }),
          cell(P(T(x.el.element_text, { bold: true }), { spacing: { after: 0 } }), ew[1]),
          cell(P(T(x.el.maps_to || "", { size: 18 }), { spacing: { after: 0 } }), ew[2]),
          cell(P([...badge({ "Not found": "Not disclosed", Framing: "Claim wording" }[x.el.status] || x.el.status, { Explicit: "green", Implicit: "green", Equivalent: "amber", "Not found": "red" }[x.el.status] || "grey"), ...(x.el.verify ? badge("Verify", "amber") : [])], { spacing: { after: 0 } }), ew[3]),
          cell(P(T([...new Set(arr(x.el.evidence).map(v => "Excerpt " + v.excerpt_number))].join(", ") || "—", { size: 17 }), { spacing: { after: 0 } }), ew[4]) ] }))]));
      }
      // caveats
      const verify = arr(dec.Verify), nf = colored.filter(x => x.el.status === "Not found");
      const cp = (dec.Construction && dec.Construction.claim_words) ? dec.Construction : null;
      if (verify.length || nf.length || cp) s2.push(table(PGL.W, [PGL.W], [new TableRow({ children: [cell([
        ...(cp ? [P(T(`Construction point · “${cp.claim_words}”`, { bold: true, color: "8A5A00" })),
                  ...(cp.question ? [P(T(cp.question, { bold: true, size: 18 }))] : []),
                  P([T("Broad reading (adopted): ", { bold: true, size: 18, color: "8A5A00" }), T(`${cp.adopted_reading} `, { size: 18 }), T(`→ ${cp.adopted_result}`, { bold: true, size: 18 })]),
                  P([T("Narrower reading: ", { bold: true, size: 18, color: "8A5A00" }), T(`${cp.narrower_reading} `, { size: 18 }), T(`→ ${cp.narrower_result}`, { bold: true, size: 18 })]),
                  P(T("Claims are read without the patent's description, which may settle the point.", { size: 17, color: C.mid }))] : []),
        ...(verify.length ? [P(T("To verify", { bold: true, color: "8A5A00" })), ...verify.map(v => P(T(v, { size: 18 })))] : []),
        ...(nf.length ? [P(T("Not found in the standard", { bold: true, color: "8A0000" })), ...nf.map(x => P(T(x.el.element_text, { size: 18 })))] : []) ], PGL.W,
        { shading: shade("FDF5E0"), borders: { top: noBorder, bottom: noBorder, right: noBorder, left: solidBorder("E8C96A", 12) } })] })]));
      if (s.essentiality_line) s2.push(P([T("Essentiality. ", { bold: true, color: C.navy }), T(s.essentiality_line, { color: C.mid })], { spacing: { before: 120, after: 80 } }));
      // reasoning
      const rs = [["Interpretation", an.Interpretation], ["Opinion", an.Overall_Opinion],
        ["Differences", an.Differences && !/^none\.?$/i.test(String(an.Differences).trim()) ? an.Differences : ""], ["Essentiality justification", dec.Justification]].filter(([, v]) => v);
      if (rs.length) { s2.push(new Paragraph({ children: [T("Reasoning", { bold: true, color: C.navy })], spacing: { before: 160, after: 60 } }));
        for (const [hh, v] of rs) s2.push(P([T(hh + ". ", { bold: true, size: 18, color: C.mid }), T(v, { size: 18, color: C.mid })])); }
      // cited excerpts
      const cx = ex.filter(x => cited.has(x.num)), ox = ex.filter(x => !cited.has(x.num));
      if (cx.length) s2.push(new Paragraph({ children: [T("Cited excerpts", { bold: true, color: C.navy })], spacing: { before: 160, after: 60 } }));
      for (const x of cx) s2.push(table(PGL.W, [PGL.W], [
        new TableRow({ children: [cell(P([T("Excerpt " + x.num + "   ", { bold: true, size: 17 }), T(x.ref, { size: 16, font: "Courier New", color: C.muted })], { spacing: { after: 0 } }), PGL.W, { shading: shade(C.surfaceAlt) })] }),
        new TableRow({ children: [cell([...(x.heading ? [P(T(x.heading, { bold: true, size: 18, color: C.navy }))] : []),
          ...highlightedParas(x.text, exR[x.num] || [], { font: "Courier New", size: 16 })], PGL.W)] }) ]));
      if (ox.length) s2.push(P(T("Other retrieved excerpts, not relied on: " + ox.map(x => `Excerpt ${x.num} (${x.ref})`).join("; ") + ". Full text in the HTML report.", { size: 16, color: C.muted }), { spacing: { before: 120, after: 80 } }));
    });

    // ── Section 3: appendix ──
    const s3 = [...sectionHeading("Appendix")];
    if (data.Summary) { s3.push(sub("Full conclusion")); s3.push(P(T(data.Summary, { color: C.mid }))); }
    const lim = String(data["Limitation(s)"] || "");
    if (lim) { const ll = lim.split("\n")[0].trim(), lb = lim.split("\n").slice(1).join("\n").trim();
      s3.push(sub("Limitations" + (ll ? ": " + ll : ""))); if (lb) s3.push(P(T(lb, { color: C.mid }))); }
    const M = data.Methodology || {};
    const terms = [{ label: "Reading of the claims", definition: "Claims are read on their own wording, without the patent's description, using the broadest technically sensible reading of the claim language. Where a narrower reasonable reading would change a result, the chart marks a Construction point with both readings; the description may resolve it." }, ...arr(M.disclosure_categories), ...arr(M.essentiality_tiers)];
    const metrics = Object.entries(M.universal_metrics || {});
    if (terms.length || metrics.length) {
      s3.push(sub("Key to terms"));
      const tw = [2600, PG.W - 2600];
      s3.push(table(PG.W, tw, [...terms.map(t => [t.label, t.definition]), ...metrics.map(([kk, v]) => [{ percentage_mapped: "Mapped", weighted_mapping: "Evidence strength (weighted mapping)" }[kk] || kk, v])]
        .map(([a, b]) => new TableRow({ children: [cell(P(T(a, { bold: true, size: 18 }), { spacing: { after: 0 } }), tw[0]), cell(P(T(b, { size: 18, color: C.mid }), { spacing: { after: 0 } }), tw[1])] }))));
    }
    if (restricted) s3.push(...restrictedNoticePage());
    s3.push(...disclaimerSection(), emptyPara());

    const page = (landscape) => ({ type: SectionType.NEXT_PAGE, page: { size: landscape ? { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE } : { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } });
    const doc = new Document({ sections: [
      { properties: page(false), headers: { default: makeHeader(PG.W) }, footers: { default: makeFooter() }, children: s1 },
      { properties: page(true),  headers: { default: makeHeader(PGL.W) }, footers: { default: makeFooter() }, children: s2 },
      { properties: page(false), headers: { default: makeHeader(PG.W) }, footers: { default: makeFooter() }, children: s3 },
    ] });
    return Packer.toBuffer(doc);
  };
};
