// Labels for the v2 claim chart (PRISM+ R1, release v1.3.0).
// A chart from the shared PRISM+ pipeline carries its standard's labels in Snapshot.labels (from the standard's
// framework): the class words and badge kinds, the implementation sides, the claim types, the ranks used for the
// questions a verdict turns on, the "who practises it" lines and the legend. Older charts don't carry them and get
// HEVC's words, exactly as before. Snapshot.tables holds, for a table with merged cells shown in an excerpt, the
// grid text the excerpt contains and the table's clean HTML, so the chart can draw the merges.
const arr = (x) => Array.isArray(x) ? x : (x == null ? [] : [x]);

// HEVC's words: the behaviour of layout v2 before R1, kept for charts without labels
const HEVC = {
  essParts(cls) {
    const [scope, side] = String(cls || '').split(' | ');
    const m = scope.match(/^Essential \((.+)\)$/);
    const label = m ? m[1].replace(/ in Main Profile$/, '').replace(/^Optional in Profile SEP$/, 'Optional (profile SEP)') : scope;
    const kind = m ? (/^Core/.test(m[1]) ? 'core' : 'opt') : (/inherent|non-technical/i.test(scope) ? 'inh' : /implementation/i.test(scope) ? 'impl' : 'no');
    return { scope, side: side || '', label, kind };
  },
  sides: { 'Encoder and Decoder': 'Encoder and decoder', 'Decoder': 'Decoder only', 'Encoder': 'Encoder only', 'Codec System': 'Codec system only' },
  categories: { Encoder: 'Encoder claim', Decoder: 'Decoder claim', Encoder_and_Decoder_System: 'Encoder and decoder claim', Other: 'Other claim type' },
  rank(t) { const x = String(t || '').toLowerCase();
    if (/not disclosed|not essential/.test(x)) return 4; if (/implementation/.test(x)) return 3;
    if (/profile/.test(x) && !/main profile/.test(x)) return 2; if (/optional/.test(x)) return 1; return 0; },
  rankLabels: ['Essential (Core)', 'Essential (Optional)', 'Essential (Profile SEP)', 'Implementation Matter', 'Not Essential'],
  resultBadge: (r) => /^\s*(essential\s*\(\s*)?core/i.test(String(r)) ? 'ess-core' : /^\s*(essential\s*\(\s*)?optional/i.test(String(r)) ? 'ess-opt' : 'ess-impl',
  whoText: (pb) => pb ? `Decoder: ${pb.decoder || ''} Encoder (${pb.encoder_case || ''}): ${pb.encoder || ''}` : '',
  legend: [{ badge: 'ess-core', label: 'Core', text: 'required in every Main / Main 10 bitstream' }, { badge: 'ess-opt', label: 'Optional', text: 'essential when an enabling flag is on' }],
  conditionPrefix: 'Only when this enabling flag is set: ',
};

function labelsFor(snap) {
  const L = (snap && snap.labels) || null;
  const tables = arr(snap && snap.tables).filter(x => x && x.grid && x.html);
  const out = { ...HEVC, standard: L ? L.standard : 'HEVC', tables };
  if (!L) return out;
  const classes = L.classes || {};
  out.essParts = (cls) => {
    const [scope, side] = String(cls || '').split(' | ');
    const c = classes[scope];
    if (!c) return { ...HEVC.essParts(cls), side: side || '' };
    return { scope, side: side || '', label: c.short || scope, kind: c.kind || 'no' };
  };
  if (L.sides) out.sides = L.sides;
  if (L.categories) out.categories = L.categories;
  if (L.ranks && Array.isArray(L.ranks.tests)) {
    const tests = L.ranks.tests.map(k => ({ rank: k.rank, re: new RegExp(k.match), unless: k.unless ? new RegExp(k.unless) : null }));
    out.rank = (t) => { const x = String(t || '').toLowerCase(); for (const k of tests) if (k.re.test(x) && !(k.unless && k.unless.test(x))) return k.rank; return 0; };
    out.rankLabels = L.ranks.short_labels || L.ranks.labels || HEVC.rankLabels;
  }
  if (Array.isArray(L.result_badges)) { const rb = L.result_badges.map(([re, b]) => [new RegExp(re, 'i'), b]);
    out.resultBadge = (r) => { for (const [re, b] of rb) if (re.test(String(r))) return b; return 'ess-impl'; }; }
  if (Array.isArray(L.practised_by) && L.practised_by.length)
    out.whoText = (pb) => pb ? L.practised_by.map(x => x.case ? `${x.label} (${pb[x.case] || ''}): ${pb[x.key] || ''}` : `${x.label}: ${pb[x.key] || ''}`).join(' ') : '';
  if (Array.isArray(L.legend)) out.legend = L.legend;
  if (typeof L.condition_prefix === 'string') out.conditionPrefix = L.condition_prefix;
  return out;
}

// A table block of an excerpt (its text exactly, as the grid of pipe rows) → the merged table, or null.
// Returns rows of cells { text, rs, cs, th, cover: [[r, c] …] } where cover lists the grid cells a merged cell spans,
// so highlights found in the grid can be placed on the merged cell.
function mergedTable(blockText, tables) {
  const norm = (s) => String(s).split('\n').map(l => l.trim()).join('\n').trim();
  const hit = arr(tables).find(x => norm(x.grid) === norm(blockText));
  if (!hit) return null;
  const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
  const rows = [];
  for (const rm of String(hit.html).matchAll(/<tr>([\s\S]*?)<\/tr>/g)) {
    const cells = [];
    for (const cm of rm[1].matchAll(/<(t[dh])((?:\s+(?:rowspan|colspan)="\d+")*)>([\s\S]*?)<\/\1>/g)) {
      const num = (n) => { const m = cm[2].match(new RegExp(`${n}="(\\d+)"`)); return m ? +m[1] : 1; };
      cells.push({ th: cm[1] === 'th', rs: num('rowspan'), cs: num('colspan'), text: unesc(cm[3]), cover: [] });
    }
    rows.push(cells);
  }
  const occ = [];
  rows.forEach((cells, r) => {
    occ[r] = occ[r] || [];
    let c = 0;
    for (const cell of cells) {
      while (occ[r][c]) c++;
      for (let dr = 0; dr < cell.rs && r + dr < rows.length; dr++) for (let dc = 0; dc < cell.cs; dc++) {
        occ[r + dr] = occ[r + dr] || []; occ[r + dr][c + dc] = true; cell.cover.push([r + dr, c + dc]);
      }
      c += cell.cs;
    }
  });
  return rows.length ? { rows } : null;
}

module.exports = { labelsFor, mergedTable, HEVC };
