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
const { locate, parseExcerptV2, excerptBlocks, latexLine, glanceV3, KEY_V3 } = require("./layout-v2");

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
  // Excerpt body (14 Oct 2026): text blocks as paragraphs with equations as readable text, tables as Word tables.
  // An evidence highlight that falls in a table shades the cells it covers.
  function excerptParas(text, ranges, base) {
    const rs = arr(ranges).filter(Boolean), out = [];
    for (const b of excerptBlocks(text)) {
      if (b.type === "text") {
        if (!b.text.trim()) continue;
        const sh = rs.filter(r => r.s < b.e && b.s < r.e).map(r => ({ ...r, s: Math.max(r.s, b.s) - b.s, e: Math.min(r.e, b.e) - b.s }));
        out.push(...highlightedParas(b.text, sh, base, {}, true));
        continue;
      }
      const ncol = Math.max(1, ...b.rows.map(r => r.length));
      const W = PGL.W - 400, w = Math.floor(W / ncol), seen = new Set();
      const line = solidBorder(C.rule, 4);
      out.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: Array(ncol).fill(w),
        rows: b.rows.map((r, ri) => new TableRow({ children: r.concat(Array(ncol - r.length).fill({ text: "", spans: [] })).map(c => {
          const hit = rs.find(x => c.spans.some(([s, e]) => s < x.e && x.s < e));
          const runs = [T(latexLine(c.text), { size: 15, bold: ri === 0, ...(hit ? { shading: { type: ShadingType.CLEAR, fill: EL[hit.k], color: "auto" } } : {}) })];
          if (hit && !seen.has(hit)) { runs.push(T(String(hit.n), { size: 15, superScript: true, bold: true, color: ELN[hit.k] })); seen.add(hit); }
          return new TableCell({ width: { size: w, type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 80, right: 80 },
            borders: { top: line, bottom: line, left: line, right: line }, ...(ri === 0 ? { shading: shade(C.surfaceAlt) } : {}),
            children: [new Paragraph({ children: runs, spacing: { after: 0 } })] });
        }) })) }));
      out.push(new Paragraph({ children: [], spacing: { after: 60 } }));
    }
    return out;
  }

  function highlightedParas(text, ranges, base, paraOpts = {}, math = false) {
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
        if (ln) runs.push(T(math ? latexLine(ln) : ln, sg.k === undefined ? base : { ...base, shading: { type: ShadingType.CLEAR, fill: EL[sg.k], color: "auto" } }));
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
    // decisive construction points: from the snapshot (b31+), else computed from the charts
    const RANK = (t) => { const x = String(t || "").toLowerCase();
      if (/not disclosed|not essential/.test(x)) return 4; if (/implementation/.test(x)) return 3;
      if (/profile/.test(x) && !/main profile/.test(x)) return 2; if (/optional/.test(x)) return 1; return 0; };
    const RLABEL = ["Essential (Core)", "Essential (Optional)", "Essential (Profile SEP)", "Implementation Matter", "Not Essential"];
    const fr = charts.map(c => { const d = c.Decision || {}; const inh = /inherent/i.test(String(d.Essentiality_Classification || ""));
      const own = inh ? 0 : (d.Disclosure === "Not Disclosed" ? 4 : RANK(d.Essentiality_Classification));
      const cp = (d.Construction && d.Construction.claim_words) ? d.Construction : null;
      const altT = cp ? String(cp.narrower_result || "") : "";
      const alt = cp && !inh ? (/not disclosed/i.test(altT) ? 4 : /disclosed/i.test(altT) ? (own === 4 ? 0 : own) : RANK(altT)) : own;
      return { idx: String(c?.Claim_Feature?.Index ?? ""), own, alt, cp }; });
    const baseRank = Math.max(0, ...fr.map(f => f.own));
    const decisive = new Map();
    if (arr(snap.decisive).length) { for (const d of arr(snap.decisive)) decisive.set(String(d.feature), d.verdict_if_alternative || ""); }
    else for (const f of fr) if (f.cp) { const r = Math.max(f.alt, ...fr.filter(g => g !== f).map(g => g.own)); if (r !== baseRank) decisive.set(f.idx, RLABEL[r]); }
    const plainDec = {}; for (const d of arr((snap.plain || {}).decisive)) plainDec[String(d.feature)] = d;
    const plainSummary = verdict.plain_summary || (() => { const b = String(verdict.bottom_line || "").replace(/^\s*(essential,\s*(core|optional)\s+in\s+[^:]{2,40}|implementation matter|not essential)\s*:\s*/i, ""); return b ? b[0].toUpperCase() + b.slice(1) : ""; })();
    const nVerify = charts.reduce((n, c) => n + arr((c.Decision || {}).Verify).length, 0);
    const parts = [];
    if (arr(align.stated_directly).length) parts.push(`${cap(featList(align.stated_directly))} ${align.stated_directly.length > 1 ? "are" : "is"} stated directly in the standard`);
    if (arr(align.by_implication).length)  parts.push(`${featList(align.by_implication)} ${align.by_implication.length > 1 ? "follow" : "follows"} by necessary implication`);
    if (arr(align.by_equivalence).length)  parts.push(`${featList(align.by_equivalence)} ${align.by_equivalence.length > 1 ? "map" : "maps"} by functional equivalence`);
    if (arr(align.not_disclosed).length)   parts.push(`${featList(align.not_disclosed)} ${align.not_disclosed.length > 1 ? "are" : "is"} not disclosed`);
    if (arr(align.inherent).length)        parts.push(`${featList(align.inherent)} ${align.inherent.length > 1 ? "are inherent components" : "is an inherent component"}`);
    const evidenceLine = parts.length ? cap(parts.join("; ")) + (nVerify ? `; ${nVerify} point${nVerify > 1 ? "s" : ""} to verify` : "") + "." : "";
    const gapsPlain = arr((snap.plain || {}).gaps).length ? arr(snap.plain.gaps)
      : charts.filter(c => (c.Decision || {}).Disclosure === "Not Disclosed").map(c => { const i = String(c?.Claim_Feature?.Index ?? ""); return { feature: i, text: (sumByIdx[i] || {}).gap || "Not found in the standard." }; });
    const band = (() => { const n = Number(verdict.weighted_percentage_mapped); return Number.isFinite(n) ? (n >= 85 ? "High" : n >= 70 ? "Moderate" : "Low") : "—"; })();
    // ── v3 (HEVC-SNAP v3.0+; D20): outcomes with confidence bands, reading questions, two answers per feature ──
    const oc = (snap.outcomes && snap.outcomes.most_likely) ? snap.outcomes : null;
    const q3 = arr(snap.questions);
    const nReaders = (q) => ((q.votes || {}).M || 0) + ((q.votes || {}).N || 0) + ((q.votes || {}).tie || 0);
    const statusWords = (q) => { const v = q.votes || {}, a = Math.max(v.M || 0, v.N || 0), b = Math.min(v.M || 0, v.N || 0);
      if (q.dropped_by_sceptic) return { notation: "a naming difference, not a dispute", case_by_case: "the same result in every case, not a dispute", contradiction: "the other reading contradicts the claim" }[q.dropped_by_sceptic] || "not a genuine dispute";
      if (q.status === "agreed") return "settled";
      if (q.status === "leans") return "one reading is clearly stronger";
      if (q.status === "split") return "both readings are reasonable";
      return "not assessed"; };
    const BUCKET = { required: ["in the standard and required", "green"], optional: ["in the standard, optional for some products", "green"], implementation: ["in the standard, an implementation choice", "amber"],
      not_in_standard: ["not found in the standard", "red"], outside: ["outside the standard", "grey"], framing: ["framing (e.g. processor, memory)", "grey"] };
    const IN_STD_KIND = { "Stated": "green", "Follows necessarily": "green", "Equivalent": "amber", "Not found": "red", "Outside the standard": "grey", "Framing": "grey" };
    const lowerFirst = (s) => String(s || "").replace(/^Essential · /, "essential · ").replace(/^Not essential/, "not essential").replace(/^Implementation matter/, "implementation matter");
    const panel = oc ? (() => {
      // At a glance (D, 15 Oct): the same text as the HTML layout, from glanceV3
      const g = glanceV3(data);
      const col = g.toss_up ? "8A5A00" : /^Not essential/.test(g.label) ? "8A0000" : /^Implementation/.test(g.label) ? "8A5A00" : "1A6B4A";
      const row = (k, words, t, feature, kcol) => P([T(k.toUpperCase() + "  ", { bold: true, size: 15, color: kcol || "1A6B4A" }),
        ...(words ? [T(`“${words}”: `, { size: 18, bold: true })] : []), T(t, { size: 18 }), ...(feature ? [T(`  (feature ${feature})`, { size: 16, color: C.mid })] : [])], { spacing: { after: 40 } });
      const out = [P(T(g.label, { font: "Georgia", size: 30, bold: true, color: col }), { spacing: { after: 40 } }),
        P([T("Most likely outcome  ·  Confidence: ", { size: 17, bold: true, color: C.mid }), T(` ${g.band} `, { size: 17, bold: true, color: "FFFFFF", shading: { type: "clear", color: "auto", fill: g.band === "High" ? "1A6B4A" : g.band === "Moderate" ? "2E7D5B" : "8A5A00" } })], { spacing: { after: 100 } }),
        ...g.arguable.map(a => P([T("ARGUABLE OUTCOME  ", { bold: true, size: 15, color: C.mid }), T(a, { font: "Georgia", size: 24, bold: true, color: "8A5A00" })], { spacing: { after: 80 } })),
        ...(plainSummary ? [P(T(plainSummary, { size: 21 }), { spacing: { after: 80 } })] : [])];
      const gh = (t) => P(T(t.toUpperCase(), { bold: true, size: 14, color: C.mid }), { spacing: { before: 100, after: 30 } });
      if (g.who.length) { out.push(gh("Who practises the claim")); for (const x of g.who) out.push(row(x.k, "", x.t)); }
      out.push(gh("What could change the result"));
      if (g.change.length) for (const x of g.change) out.push(row(x.k, x.words, x.t, x.feature, "8A5A00"));
      else out.push(P(T(g.nothing, { size: 18 }), { spacing: { after: 40 } }));
      if (g.notRequired.length) { out.push(gh("Not required by the standard")); for (const x of g.notRequired) out.push(row(`Feature ${x.feature}`, "", x.t, "", "8A5A00")); }
      if (gapsPlain.length) {
        out.push(P(T(`What is missing (${gapsPlain.length})`, { bold: true, size: 19, color: "8A0000" }), { spacing: { before: 80, after: 40 } }));
        for (const gp of gapsPlain) out.push(P([T(`Feature ${gp.feature}  `, { bold: true, size: 18 }), T(gp.text, { size: 18 })], { spacing: { after: 40 } }));
      }
      if (g.points.length) { out.push(gh("Construction points settled")); for (const x of g.points) out.push(P([T(`“${x.words}”: `, { bold: true, size: 16 }), T(x.t, { size: 16, color: C.mid })], { spacing: { after: 20 } })); }
      return out;
    })() : [
      P([T(vLabel, { font: "Georgia", size: 30, bold: true, color: vText }), T(`     ${verdict.features_aligned ?? "—"}/${verdict.features_technical ?? "—"} features aligned  ·  Evidence strength: ${band}`, { size: 17, color: C.mid })], { spacing: { after: 100 } }),
      ...(plainSummary ? [P(T(plainSummary, { size: 21 }), { spacing: { after: 80 } })] : []),
      ...(evidenceLine ? [P(T(evidenceLine, { size: 17, color: C.mid }), { spacing: { after: 80 } })] : []),
    ];
    if (gapsPlain.length && !oc) {
      panel.push(P(T(`What is missing (${gapsPlain.length})`, { bold: true, size: 19, color: "8A0000" }), { spacing: { before: 80, after: 40 } }));
      for (const g of gapsPlain) panel.push(P([T(`Feature ${g.feature}  `, { bold: true, size: 18 }), T(g.text, { size: 18 })], { spacing: { after: 40 } }));
    }
    if (decisive.size && !oc) {
      panel.push(P(T(`The question${decisive.size > 1 ? "s" : ""} the verdict turns on (${decisive.size})`, { bold: true, size: 19, color: "8A5A00" }), { spacing: { before: 100, after: 40 } }));
      for (const f of fr.filter(f => decisive.has(f.idx) && f.cp)) {
        const p = plainDec[f.idx] || {};
        panel.push(P([T(`Feature ${f.idx}  `, { bold: true, size: 18 }), T(p.question || f.cp.question || `How should “${f.cp.claim_words}” be read?`, { size: 18 }), ...(f.cp.close_call === true ? [T("   (close call)", { size: 16, color: C.mid })] : [])], { spacing: { after: 20 } }));
        panel.push(P([T("If yes ", { bold: true, size: 17 }), T(`(this chart's reading): ${p.if_adopted || ("the claim is " + String(verdict.classification || "").toLowerCase())}    `, { size: 17, color: C.mid }),
                       T("If no: ", { bold: true, size: 17 }), T(p.if_alternative || `the claim would be ${decisive.get(f.idx)}`, { size: 17, color: C.mid })], { spacing: { after: 60 } }));
      }
    }
    const jr = snap.joint_route || null;
    if (jr && arr(jr.features).length && !oc) panel.push(P([T("Arguable route: ", { bold: true, size: 18, color: "8A5A00" }), T(((snap.plain || {}).route_note) || `if features ${arr(jr.features).join(" and ")} were all read the other way, the claim would be ${String(jr.verdict || "").toLowerCase()}.`, { size: 18 })], { spacing: { before: 80, after: 40 } }));
    s1.push(table(PG.W, [PG.W], [new TableRow({ children: [cell(panel, PG.W, { shading: shade(oc ? (oc.toss_up ? "FDF5E0" : /^Not essential/.test(oc.most_likely.side_label) ? "FDF0F0" : /^Implementation/.test(oc.most_likely.side_label) ? "FDF5E0" : "EAF5EF") : vFill) })] })]));

    // Summary mapping
    s1.push(...sectionHeading("Summary mapping"));
    s1.push(P(T("One row per claim feature, in the claim's own words. The full analysis and evidence for each is in the detailed claim chart.", { color: C.mid, size: 18 })));
    const rest = PG.W - 450;
    const sw = [450, Math.round(rest * 0.31), Math.round(rest * 0.23), Math.round(rest * 0.14)]; sw.push(PG.W - sw.reduce((a, b) => a + b, 0));
    const caveatOf = (gap, disc) => gap && gap !== "None" ? gap : disc === "Not Disclosed" ? "Not disclosed in the standard" : disc === "Functionally Equivalent" ? "Relies on functional equivalence" : "";
    if (oc) {
      const sw3 = [450, Math.round(rest * 0.27), Math.round(rest * 0.2), Math.round(rest * 0.14), Math.round(rest * 0.16)]; sw3.push(PG.W - sw3.reduce((a, b) => a + b, 0));
      s1.push(table(PG.W, sw3, [headRow(["#", "Claim feature", "Maps to", "In the standard?", "For products", "Gap or caveat"], sw3), ...charts.map(c => {
        const i = String(c?.Claim_Feature?.Index ?? ""), s = sumByIdx[i] || {}, dec = c.Decision || {}, a = s.answers || {};
        const cav = caveatOf(s.gap, dec.Disclosure), flags = arr(dec.Gating_Flags);
        const qs = arr(s.questions).map(qid => q3.find(q => q.qid === qid)).filter(Boolean);
        return new TableRow({ cantSplit: true, children: [
          cell(P(T(i, { bold: true }), { spacing: { after: 0 } }), sw3[0]),
          cell(P(T(c?.Claim_Feature?.Text || s.short_label || "", { size: 18 }), { spacing: { after: 0 } }), sw3[1]),
          cell(P(T(s.maps_to || "", { size: 17, color: C.mid }), { spacing: { after: 0 } }), sw3[2]),
          cell([P(badge(a.in_standard || dec.Disclosure || "", IN_STD_KIND[a.in_standard] || "grey"), { spacing: { after: 0 } })], sw3[3]),
          cell([P(badge(a.for_products || "", (BUCKET[s.bucket] || ["", "grey"])[1]), { spacing: { after: 0 } })], sw3[4]),
          cell([...(cav || (!qs.length && !flags.length) ? [P(T(cav || "None", { size: 17, color: cav ? "8A5A00" : C.muted }), { spacing: { after: 0 } })] : []),
                ...(flags.length ? [P([...badge("Encoder switch", "green"), T(" (" + flags.join(" / ") + ")", { size: 14, font: "Courier New", color: C.mid })], { spacing: { after: 0 } })] : []),
                ...qs.map(q => P([...badge(q.decides ? "Decisive point" : "Construction point", "amber"), T(` “${q.claim_words}”: ${statusWords(q)}`, { size: 16 })], { spacing: { after: 0 } }))], sw3[5]) ] });
      })]));
    } else
    s1.push(table(PG.W, sw, [headRow(["#", "Claim feature", "Maps to", "Essentiality", "Gap or caveat"], sw), ...charts.map(c => {
      const i = String(c?.Claim_Feature?.Index ?? ""), s = sumByIdx[i] || {}, dec = c.Decision || {}, e = essParts(dec.Essentiality_Classification);
      const flags = arr(dec.Gating_Flags);
      const cpS = (dec.Construction && dec.Construction.claim_words && decisive.has(i)) ? dec.Construction : null;
      const tools = [...new Set(arr(dec.Condition_Tools).map(x => x && x.tool).filter(Boolean))];
      const cav = caveatOf(s.gap, dec.Disclosure);
      return new TableRow({ cantSplit: true, children: [
        cell(P(T(i, { bold: true }), { spacing: { after: 0 } }), sw[0]),
        cell(P(T(c?.Claim_Feature?.Text || s.short_label || "", { size: 18 }), { spacing: { after: 0 } }), sw[1]),
        cell(P(T(s.maps_to || "", { size: 17, color: C.mid }), { spacing: { after: 0 } }), sw[2]),
        cell([P(badge(e.label, e.kind), { spacing: { after: 0 } })], sw[3]),
        cell([...(cav || (!cpS && !flags.length) ? [P(T(cav || "None", { size: 17, color: cav ? "8A5A00" : C.muted }), { spacing: { after: 0 } })] : []),
              ...(flags.length ? [P([...badge("Condition", "green"), T(" Needs " + (tools.length ? tools.join(" and ") + " " : ""), { size: 16, color: C.mid }), T("(" + flags.join(" / ") + ")", { size: 14, font: "Courier New", color: C.mid })], { spacing: { after: 0 } })] : []),
              ...(cpS ? [P([...badge("Decisive", "amber"), T(" " + ((plainDec[i] || {}).question || cpS.question || `How should “${cpS.claim_words}” be read?`), { size: 16 })], { spacing: { after: 0 } })] : [])], sw[4]) ] });
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
        cell([...highlightedParas(ftext, ftR, { size: 22 }), oc && s.answers
          ? P([T("In the standard? ", { size: 16, bold: true, color: C.muted }), ...badge(s.answers.in_standard, IN_STD_KIND[s.answers.in_standard] || "grey"),
               T("   For products: ", { size: 16, bold: true, color: C.muted }), ...badge(s.answers.for_products, (BUCKET[s.bucket] || ["", "grey"])[1]),
               ...(arr(dec.Gating_Flags).length ? [T("   " + arr(dec.Gating_Flags).join(" / "), { size: 15, font: "Courier New", color: C.mid })] : [])], { spacing: { before: 80, after: 0 } })
          : P([...badge(dec.Disclosure || "", discKind(dec.Disclosure)), ...badge(e.label, e.kind),
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
          cell(P([...badge({ "Not found": "Not disclosed", Framing: "Claim wording", Context: "Outside the standard" }[x.el.status] || x.el.status, { Explicit: "green", Implicit: "green", Equivalent: "amber", "Not found": "red" }[x.el.status] || "grey"), ...(x.el.verify ? badge("Verify", "amber") : [])], { spacing: { after: 0 } }), ew[3]),
          cell(P(T([...new Set(arr(x.el.evidence).map(v => "Excerpt " + v.excerpt_number))].join(", ") || "—", { size: 17 }), { spacing: { after: 0 } }), ew[4]) ] }))]));
      }
      // v3: reading questions, both routes (the chart's reading is used throughout; the other is a separate route)
      if (oc) for (const r of arr((c.Routes || {}).questions)) {
        const q = q3.find(x => x.qid === r.qid) || r;
        const used = r.reading_used || q.leading, other = used === "M" ? "N" : "M";
        const alt = arr((c.Routes || {}).alternatives).find(x => x.qid === r.qid);
        const altRes = alt && alt.analysis ? (alt.analysis.Disclosure_Decision === "Not Disclosed" ? "Not disclosed" : `${alt.analysis.Disclosure_Decision || ""}${alt.classification ? " · " + alt.classification.classification : ""}`) : "Not analysed: the point is settled";
        const half = Math.floor((PGL.W - 480) / 2);
        const col = (label, reading, result, adopted) => cell([
          P([T(label.toUpperCase(), { bold: true, size: 15, color: C.mid }), ...(adopted ? [T("   used throughout", { bold: true, size: 15, color: "1A6B4A" })] : [])], { spacing: { after: 60 } }),
          P(T(reading, { size: 18 }), { spacing: { after: 60 } }),
          P([T("→ ", { size: 18, color: C.mid }), T(result, { bold: true, size: 18, color: adopted ? "1A6B4A" : "8A5A00" })], { spacing: { after: 0 } })],
          half, { shading: shade("FFFFFF"), borders: { top: solidBorder(adopted ? "9FD1B6" : "E8C96A", 6), bottom: solidBorder(adopted ? "9FD1B6" : "E8C96A", 6), left: solidBorder(adopted ? "1A6B4A" : "E8C96A", adopted ? 18 : 6), right: solidBorder(adopted ? "9FD1B6" : "E8C96A", 6) } });
        s2.push(table(PGL.W, [PGL.W], [new TableRow({ children: [cell([
          P(T(`Construction point · “${q.claim_words}”`, { bold: true, color: "8A5A00" })),
          table(half * 2, [half, half], [new TableRow({ children: [col("This chart's reading", used === "M" ? q.reading_M : q.reading_N, dec.Disclosure || "", true), col("Other reading", other === "M" ? q.reading_M : q.reading_N, altRes, false)] })]),
          P(T(`${cap(statusWords(q))}. Claims are read without the patent's description, which may settle the point.`, { size: 17, color: C.mid }), { spacing: { before: 80 } })],
          PGL.W, { shading: shade("FDF5E0") })] })]));
      }
      // caveats
      const verify = arr(dec.Verify), nf = colored.filter(x => x.el.status === "Not found"), ctx = colored.filter(x => x.el.status === "Context");
      const cp = (dec.Construction && dec.Construction.claim_words && decisive.has(i)) ? dec.Construction : null;
      if (verify.length || nf.length || ctx.length || cp) s2.push(table(PGL.W, [PGL.W], [new TableRow({ children: [cell([
        ...(cp ? [P(T(`Decisive question · “${cp.claim_words}”`, { bold: true, color: "8A5A00" })),
                  ...(cp.question ? [P(T(cp.question, { bold: true, size: 18 }))] : []),
                  (() => { const half = Math.floor((PGL.W - 480) / 2);
                    const col = (label, reading, result, adopted) => cell([
                      P([T(label.toUpperCase(), { bold: true, size: 15, color: C.mid }), ...(adopted ? [T("   adopted", { bold: true, size: 15, color: "1A6B4A" })] : [])], { spacing: { after: 60 } }),
                      P(T(reading, { size: 18 }), { spacing: { after: 60 } }),
                      P([T("→ ", { size: 18, color: C.mid }), T(result, { bold: true, size: 18, color: adopted ? "1A6B4A" : "8A5A00" })], { spacing: { after: 0 } })],
                      half, { shading: shade("FFFFFF"), borders: { top: solidBorder(adopted ? "9FD1B6" : "E8C96A", 6), bottom: solidBorder(adopted ? "9FD1B6" : "E8C96A", 6), left: solidBorder(adopted ? "1A6B4A" : "E8C96A", adopted ? 18 : 6), right: solidBorder(adopted ? "9FD1B6" : "E8C96A", 6) } });
                    return table(half * 2, [half, half], [new TableRow({ children: [col("This chart's reading", cp.adopted_reading, cp.adopted_result, true), col("Arguable route", cp.narrower_reading, cp.narrower_result, false)] })]); })(),
                  ...(cp.reason ? [P(T(cp.reason + (cp.close_call === true ? " A close call." : ""), { size: 17, color: C.mid }), { spacing: { before: 80 } })] : []),
                  P(T("Claims are read without the patent's description, which may settle the point.", { size: 17, color: C.mid }), { spacing: { before: 40 } })] : []),
        ...(verify.length ? [P(T("To verify", { bold: true, color: "8A5A00" })), ...verify.map(v => P(T(v, { size: 18 })))] : []),
        ...(nf.length ? [P(T("Not found in the standard", { bold: true, color: "8A0000" })), ...nf.map(x => P(T(x.el.element_text, { size: 18 })))] : []),
        ...(ctx.length ? [P(T("Outside the standard (application or product setting; not required by the standard)", { bold: true, color: C.mid })), ...ctx.map(x => P(T(x.el.element_text, { size: 18 })))] : []) ], PGL.W,
        { shading: shade("FDF5E0"), borders: { top: noBorder, bottom: noBorder, right: noBorder, left: solidBorder("E8C96A", 12) } })] })]));
      // reasoning
      const pb = dec.Practised_By || null;
      const rs = [["Who practises it", pb ? `Decoder: ${pb.decoder || ""} Encoder (${pb.encoder_case || ""}): ${pb.encoder || ""}` : ""], ["Interpretation", an.Interpretation], ["Opinion", an.Overall_Opinion],
        ["Differences", an.Differences && !/^none\.?$/i.test(String(an.Differences).trim()) ? an.Differences : ""], ["Essentiality justification", dec.Justification]].filter(([, v]) => v);
      if (rs.length) { s2.push(new Paragraph({ children: [T("Reasoning", { bold: true, color: C.navy })], spacing: { before: 160, after: 60 } }));
        for (const [hh, v] of rs) s2.push(P([T(hh + ". ", { bold: true, size: 18, color: C.mid }), T(v, { size: 18, color: C.mid })])); }
      // cited excerpts
      const cx = ex.filter(x => cited.has(x.num)), ox = ex.filter(x => !cited.has(x.num));
      if (cx.length) s2.push(new Paragraph({ children: [T("Cited excerpts", { bold: true, color: C.navy })], spacing: { before: 160, after: 60 } }));
      for (const x of cx) s2.push(table(PGL.W, [PGL.W], [
        new TableRow({ children: [cell(P([T("Excerpt " + x.num + "   ", { bold: true, size: 17 }), T(x.ref, { size: 16, font: "Courier New", color: C.muted })], { spacing: { after: 0 } }), PGL.W, { shading: shade(C.surfaceAlt) })] }),
        new TableRow({ children: [cell([...(x.heading ? [P(T(x.heading, { bold: true, size: 18, color: C.navy }))] : []),
          ...excerptParas(x.text, exR[x.num] || [], { font: "Courier New", size: 16 })], PGL.W)] }) ]));
      if (ox.length) s2.push(P(T("Other retrieved excerpts, not relied on: " + ox.map(x => `Excerpt ${x.num} (${x.ref})`).join("; ") + ". Full text in the HTML report.", { size: 16, color: C.muted }), { spacing: { before: 120, after: 80 } }));
    });

    // ── Section 3: appendix ──
    const s3 = [...sectionHeading("Appendix")];
    if (data.Summary) { s3.push(sub("Full conclusion")); s3.push(P(T(data.Summary, { color: C.mid }))); }
    const lim = String(data["Limitation(s)"] || "");
    if (lim) { const ll = lim.split("\n")[0].trim(), lb = lim.split("\n").slice(1).join("\n").trim();
      s3.push(sub("Limitations" + (ll ? ": " + ll : ""))); if (lb) s3.push(P(T(lb, { color: C.mid }))); }
    const M = data.Methodology || {};
    const terms3 = KEY_V3;
    // v3: the methodology's essentiality tiers describe the enabling-flag test, which D19 replaced
    const terms = oc ? [...terms3, ...arr(M.disclosure_categories)] : [{ label: "Evidence strength", definition: "High where most features are stated directly in the standard; Moderate where several rest on necessary implication, a point to verify, or equivalence; Low where the mapping depends substantially on these or features are not disclosed. (Weighted mapping of at least 85%, 70–84%, below 70%.)" }, { label: "Reading of the claims", definition: "Claims are read on their own wording, without the patent's description. Where claim words can reasonably be read more than one way, the chart adopts the reading a skilled reader would find more natural, as a judge would, and shows the other as an arguable route with its result. Where the two are genuinely balanced, the chart adopts the reading under which the claim maps and marks a close call." }, ...arr(M.disclosure_categories), ...arr(M.essentiality_tiers)];
    const ckey = (id) => String(id).split(".").map(p => /^\d+$/.test(p) ? p.padStart(4, "0") : p).join(".");
    const cl = arr(align.provisions).filter(p => /^([0-9]+|[A-H])(\.[0-9]+)+$/.test(String(p.clause || ""))).sort((a, b) => ckey(a.clause) < ckey(b.clause) ? -1 : 1);
    if (cl.length) {
      s3.push(sub("Clauses relied on"));
      const cw = [1500, PG.W - 1500 - 1700, 1700];
      s3.push(table(PG.W, cw, cl.map(p => new TableRow({ children: [
        cell(P(T(p.clause, { bold: true, font: "Courier New", size: 17 }), { spacing: { after: 0 } }), cw[0]),
        cell(P(T(String(p.title || "").replace(/\s*\((full clause|clause extract)\)\s*$/i, ""), { size: 18 }), { spacing: { after: 0 } }), cw[1]),
        cell(P(T(arr(p.features).map(f => "Feature " + f).join(", "), { size: 16, color: C.mid }), { spacing: { after: 0 } }), cw[2]) ] }))));
    }
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
