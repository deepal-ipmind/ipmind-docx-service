"use strict";
// ─────────────────────────────────────────────────────────────────────────────
// Claim chart layout v2 (HTML) — ipmind-docx-service
// Rendered when /generate-html is called with ?layout=v2 and the payload
// carries a Snapshot (HEVC workflow, batch 10 onwards). Otherwise layout v1.
//   At a glance: verdict, where it aligns (by provision), gaps and conditions
//   Summary mapping: one row per feature
//   Detailed claim chart: per-element colour highlighting in the claim text,
//     the element table and the cited excerpts
//   Appendix: full conclusion, limitations, key to terms, disclaimer
// ─────────────────────────────────────────────────────────────────────────────
// Shared with layout v1 (copied verbatim from server.js)
const LOGO_SVG = "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" height=\"36\" viewBox=\"0 0 3144.8497854077254 1027.5281652360513\"><g transform=\"scale(7.242489270386266) translate(10, 10)\"><defs id=\"SvgjsDefs1027\"/><g id=\"SvgjsG1028\" featureKey=\"symbolGroupContainer\" transform=\"matrix(1.16515289568328,0,0,1.16515289568328,0.000007264315552150539,0.000007264315552150539)\" fill=\"#fff\"><path d=\"M52.3 104.6a52.3 52.3 0 1 1 52.3-52.3 52.4 52.4 0 0 1-52.3 52.3zm0-102.3a50 50 0 1 0 50 50 50 50 0 0 0-50-50z\"/></g><g id=\"SvgjsG1029\" featureKey=\"2ou6gm-0\" transform=\"matrix(0.9971509971509972,0,0,0.9971509971509972,264.8062678062678,-335.4786324786325)\" fill=\"#fff\"><path d=\"M-167.5,390.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-165.5,389.6-166.4,390.5-167.5,390.5z M-177.5,428.5c-2.2,0-4-1.8-4-4s1.8-4,4-4c2.2,0,4,1.8,4,4S-175.3,428.5-177.5,428.5z M-177.5,410.5c-2.2,0-4-1.8-4-4s1.8-4,4-4c2.2,0,4,1.8,4,4S-175.3,410.5-177.5,410.5z M-177.5,392.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-173.5,390.7-175.3,392.5-177.5,392.5z M-177.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-173.5,372.7-175.3,374.5-177.5,374.5z M-194.5,414.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7c3.9,0,7,3.1,7,7C-187.5,411.4-190.6,414.5-194.5,414.5z M-194.5,394.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7c3.9,0,7,3.1,7,7C-187.5,391.4-190.6,394.5-194.5,394.5z M-195.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-191.5,372.7-193.3,374.5-195.5,374.5z M-195.5,362.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-193.5,361.6-194.4,362.5-195.5,362.5z M-214.5,414.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7s7,3.1,7,7C-207.5,411.4-210.6,414.5-214.5,414.5z M-214.5,394.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7s7,3.1,7,7C-207.5,391.4-210.6,394.5-214.5,394.5z M-213.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-209.5,372.7-211.3,374.5-213.5,374.5z M-213.5,362.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-211.5,361.6-212.4,362.5-213.5,362.5z M-231.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-227.5,372.7-229.3,374.5-231.5,374.5z M-231.5,384.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-235.5,386.3-233.7,384.5-231.5,384.5z M-241.5,408.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-239.5,407.6-240.4,408.5-241.5,408.5z M-241.5,390.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-239.5,389.6-240.4,390.5-241.5,390.5z M-231.5,402.5c2.2,0,4,1.8,4,4s-1.8,4-4,4c-2.2,0-4-1.8-4-4S-233.7,402.5-231.5,402.5z M-231.5,420.5c2.2,0,4,1.8,4,4s-1.8,4-4,4c-2.2,0-4-1.8-4-4S-233.7,420.5-231.5,420.5z M-213.5,420.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-217.5,422.3-215.7,420.5-213.5,420.5z M-213.5,432.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-215.5,433.4-214.6,432.5-213.5,432.5z M-195.5,420.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-199.5,422.3-197.7,420.5-195.5,420.5z M-195.5,432.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-197.5,433.4-196.6,432.5-195.5,432.5z M-167.5,404.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-169.5,405.4-168.6,404.5-167.5,404.5z\" style=\"fill-rule:evenodd;clip-rule:evenodd;\"/></g><g id=\"SvgjsG1030\" featureKey=\"kZnDdN-0\" transform=\"matrix(3.8775259911441498,0,0,3.8775259911441498,137.1154802767278,2.6123700442792526)\" fill=\"#fff\"><path d=\"M2.8906 8.457 c-0.88867 0 -1.6309 -0.72266 -1.6309 -1.6211 c0 -0.88867 0.74219 -1.6113 1.6309 -1.6113 c0.86914 0 1.6113 0.72266 1.6113 1.6113 c0 0.89844 -0.74219 1.6211 -1.6113 1.6211 z M1.4551 20 l0 -10.039 l2.832 0 l0 10.039 l-2.832 0 z M13.0859875 9.766 c2.6465 0 4.834 1.9434 4.834 5.2344 s-2.1875 5.2344 -4.834 5.2344 c-1.3086 0 -2.4805 -0.50781 -3.0762 -1.4258 l0 6.0742 l-2.8125 0 l0 -14.922 l2.666 0 l0.078125 1.3477 c0.55664 -0.99609 1.7773 -1.543 3.1445 -1.543 z M12.4511875 17.9004 c1.4746 0 2.6563 -1.0742 2.6563 -2.9004 s-1.1816 -2.9004 -2.6563 -2.9004 c-1.5039 0 -2.6758 1.1426 -2.6758 2.9004 s1.1719 2.9004 2.6758 2.9004 z M37.129296875 9.766 c2.1484 0 3.5352 1.0938 3.5352 3.1543 l0 7.0801 l-2.8125 0 l0 -6.2793 c0 -1.1816 -0.74219 -1.6992 -1.582 -1.6992 c-1.0059 0 -1.8945 0.57617 -1.8945 2.3145 l0 5.6641 l-2.8418 0 l0 -6.25 c0 -1.2012 -0.72266 -1.7285 -1.6113 -1.7285 c-0.97656 0 -1.8848 0.57617 -1.8848 2.4609 l0 5.5176 l-2.8027 0 l0 -10.039 l2.8027 0 l0 1.1816 c0.66406 -0.83008 1.7871 -1.3086 3.1152 -1.3086 z M44.833959375 8.457 c-0.88867 0 -1.6309 -0.72266 -1.6309 -1.6211 c0 -0.88867 0.74219 -1.6113 1.6309 -1.6113 c0.86914 0 1.6113 0.72266 1.6113 1.6113 c0 0.89844 -0.74219 1.6211 -1.6113 1.6211 z M43.398459375 20 l0 -10.039 l2.832 0 l0 10.039 l-2.832 0 z M55.068346875 9.766 c2.4121 0 3.7402 1.25 3.7402 3.4766 l0 6.7578 l-2.8223 0 l0 -6.1523 c0 -1.3379 -0.83008 -1.8262 -1.7969 -1.8262 c-1.1621 0 -2.2168 0.58594 -2.2363 2.4414 l0 5.5371 l-2.8125 0 l0 -10.039 l2.8125 0 l0 1.1133 c0.70313 -0.83008 1.7871 -1.3086 3.1152 -1.3086 z M68.652325 5 l2.8125 0 l0 15 l-2.666 0 l-0.068359 -1.3086 c-0.57617 0.98633 -1.7871 1.543 -3.1543 1.543 c-2.6465 0 -4.834 -1.9531 -4.834 -5.2344 s2.1973 -5.2344 4.834 -5.2344 c1.3184 0 2.4805 0.49805 3.0762 1.4063 l0 -6.1719 z M66.220725 17.9004 c1.4941 0 2.6563 -1.1426 2.6563 -2.9004 s-1.1719 -2.9102 -2.6563 -2.9102 c-1.4941 0 -2.666 1.1035 -2.666 2.9102 c0 1.7969 1.1719 2.9004 2.666 2.9004 z\"/></g></g></svg>";
const DISCLAIMER_ITEMS_HTML = [
  "<strong>Preliminary and Informational Nature:</strong> The present work product was generated using a prototype AI model and is provided for informational purposes only. It does not constitute a legal or technical opinion regarding the essentiality or non-essentiality of any patent claim to any technical standard. It is not a substitute for legal or technical advice, and clients are strongly encouraged to seek independent professional counsel before relying on this material for purposes such as licensing, enforcement, or infringement analysis.",
  "<strong>Scope of Analysis:</strong> The analysis is limited to the individual patent claim(s) identified in the chart and does not take into account the full patent specification, including the description and drawings. Consequently, any interpretation of claim scope is based on the claim language alone and may differ from that reached through a full legal construction under applicable law.",
  "<strong>Referencing of Standards:</strong> Where citations to section numbers, table numbers, or figure numbers in a technical standard are provided, they are included for convenience only. While care is taken in referencing, these citations should not be relied upon as authoritative without verification against the official version of the standard.",
  "<strong>Interpretation of Standards:</strong> References to technical standards are based on publicly available documents. Where relevant, excerpts are cited in text form. Figures and diagrams from such standards are not reproduced; instead, any associated visual content is paraphrased using descriptive language. Such paraphrasing should not be construed as a verbatim or authoritative interpretation of the standard itself.",
  "<strong>Subjectivity of Essentiality:</strong> Determinations of potential alignment between a patent claim and a standard may depend on how specific terms or functional steps are construed. What may appear to correspond closely under one interpretation may be viewed as merely analogous under another. This assessment is inherently interpretive and does not reflect a consensus view or judicial determination.",
  "<strong>Implementation Considerations:</strong> The presence of a feature in a standard does not imply that all compliant implementations necessarily use that feature. A compliant product may omit or bypass specific technical elements referenced in a patent claim.",
  "<strong>Alternative Solutions:</strong> Standards may include multiple options or alternative techniques to achieve similar functionality. A given patent claim may correspond to one such option, but not to others that are also compliant with the standard.",
  "<strong>Legal Proceedings:</strong> In the context of litigation, essentiality determinations typically require a far more detailed analysis, including expert testimony, claim construction under applicable law, and examination of implementation evidence. The present assessment should not be relied upon for litigation, licensing negotiation, or investment decisions without further professional review."
];
const RESTRICTED_NOTICE = "This report is confidential and provided solely for internal use in connection " +
  "with patent licensing, portfolio evaluation, or standards-related strategy. It must " +
  "not be published, posted, or circulated to any third party without IP Mind\u2019s prior " +
  "written consent. Where disclosure to a counterparty is necessary, the report may be " +
  "shared in full or in part provided the counterparty is bound by a written " +
  "confidentiality undertaking that places equivalent restrictions on use and further " +
  "distribution, and that requires attribution of IP Mind\u2019s authorship to be retained. " +
  "The recipient must not use this report to replicate, benchmark, or train models " +
  "intended to reproduce IP Mind\u2019s methodology or outputs, or to develop competing " +
  "analysis products or services.";
const CSS_V2 = "\n:root{--brand:#ff6734;--header:#0f1f38;--heading:#0f1f38;--ink:#1c1c2e;--mid:#4a4a6a;--muted:#6b6b88;--rule:#e2e2ed;--bg:#fafaf8;--surface:#fff;--surface-alt:#f4f4f0;\n--green:#1a6b4a;--green-bg:#eaf5ef;--green-line:#9fd1b6;--amber:#8a5a00;--amber-bg:#fdf5e0;--amber-line:#e8c96a;--red:#8a0000;--red-bg:#fdf0f0;\n--e1:#cddcf5;--e1l:#2f5ea8;--e2:#e4d6f6;--e2l:#6a45a6;--e3:#c9ebe5;--e3l:#22746d;--grey:#ececf2;\n--s-ok:#d3eddf;--s-okl:#1a6b4a;--s-ver:#f8e7b8;--s-verl:#8a5a00;--s-gap:#f6d3d3;--s-gapl:#8a0000;\n--serif:'Playfair Display',Georgia,'Times New Roman',serif;--sans:'Source Sans 3','Segoe UI',Helvetica,Arial,sans-serif;--mono:'Source Code Pro',Consolas,'Courier New',monospace}\nhtml{scroll-behavior:smooth}\n*,*::before,*::after{box-sizing:border-box}\nbody{margin:0;background:var(--bg);color:var(--ink);font:400 15px/1.6 var(--sans)}\na{color:inherit} a:focus-visible,summary:focus-visible{outline:2px solid var(--brand);outline-offset:2px}\n.brand-rule{height:4px;background:var(--brand)}\n.hdr{background:var(--header)} .hdr-in{max-width:1040px;margin:0 auto;padding:22px 32px;display:flex;justify-content:space-between;align-items:center;gap:16px}\n.hdr svg{height:30px;width:auto} .conf{font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.6);border:1px solid rgba(255,255,255,.25);padding:4px 12px;border-radius:2px}\nmain{max-width:1040px;margin:0 auto;padding:0 32px 64px}\n.ident{padding:36px 0 8px} .pills{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px}\n.pill{font-size:12px;font-weight:600;padding:3px 10px;border-radius:2px;background:var(--surface);border:1px solid var(--rule);color:var(--mid)}\n.pill.std{border-color:var(--brand);color:var(--brand)}\nh1{font:700 30px/1.2 var(--serif);color:var(--heading);margin:0 0 6px}\n.owner{color:var(--muted);margin:0}\ndiv.claim{margin-top:18px;border-left:3px solid var(--brand);background:var(--surface);padding:12px 20px;border-radius:0 6px 6px 0}\n.claim-h{font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--brand)} div.claim p{white-space:pre-line;font-style:italic;color:var(--mid);margin:10px 0 4px}\nh2{font:600 22px/1.3 var(--serif);color:var(--heading);margin:52px 0 18px;display:flex;align-items:center;gap:16px}\nh2::after{content:\"\";flex:1;height:1px;background:var(--rule)}\n.verdict{background:var(--green-bg);border:1px solid var(--green-line);border-radius:12px;padding:24px 28px;display:grid;grid-template-columns:1fr auto;gap:12px 32px;align-items:center}\n.v-label{font:700 26px/1.25 var(--serif);color:var(--green);margin:0}\n.v-line{margin:6px 0 0;color:var(--ink);font-size:16px;max-width:60ch}\n.v-stats{display:flex;gap:28px} .stat b{display:block;font:700 26px/1 var(--sans);color:var(--heading)} .stat span{font-size:12.5px;color:var(--mid)}\n.boxes{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-top:16px}\n.box{background:var(--surface);border:1px solid var(--rule);border-top:3px solid var(--rule);border-radius:6px;padding:18px 20px}\n.boxes.two{grid-template-columns:3fr 2fr}\n.prov li{display:grid;grid-template-columns:88px 1fr;gap:2px 12px;align-items:baseline}\n.prov b{font-family:var(--mono);font-size:13px;color:var(--heading)}\n.prov .chips{grid-column:2;display:flex;flex-wrap:wrap;gap:6px;margin-top:4px}\na.chip{font-size:12px;font-weight:600;text-decoration:none;color:var(--heading);border:1px solid var(--rule);background:var(--surface-alt);padding:1px 8px;border-radius:10px}\na.chip:hover{border-color:var(--heading)}\n.box h3{font:600 15px/1.3 var(--sans);color:var(--heading);margin:0 0 10px}\n.box.align{border-top-color:var(--green-line)} .box.gaps{border-top-color:var(--amber-line)} .box.prov{border-top-color:var(--brand)}\n.box ul{list-style:none;margin:0;padding:0} .box li{padding:7px 0;border-top:1px solid var(--rule);font-size:14px} .box li:first-child{border-top:0;padding-top:0}\n.box .lead{font-size:14px;margin:0 0 10px;color:var(--mid)}\n.tag{display:inline-block;font-size:11px;font-weight:700;padding:1px 7px;border-radius:2px;margin-right:6px}\n.tag.cond{background:var(--green-bg);color:var(--green)} .tag.ver{background:var(--amber-bg);color:var(--amber)} .tag.gap{background:var(--red-bg);color:var(--red)}\n.fref{color:var(--muted);font-size:13px;white-space:nowrap}\ncode,.mono{font-family:var(--mono);font-size:.88em}\n.tablewrap{overflow-x:auto;border:1px solid var(--rule);border-radius:6px;background:var(--surface)}\ntable{border-collapse:collapse;width:100%}\n.align-t th,.align-t td{text-align:left;vertical-align:top;padding:12px 14px;border-top:1px solid var(--rule);font-size:14px}\n.align-t thead th{border-top:0;font-size:12.5px;font-weight:600;color:var(--muted);background:var(--surface-alt)}\n.align-t th[scope=row] a{display:inline-grid;place-items:center;width:26px;height:26px;border-radius:50%;background:var(--header);color:#fff;font-weight:700;font-size:13px;text-decoration:none}\n.align-t .maps{color:var(--mid)} .none{color:var(--muted)}\n.badge{display:inline-block;font-size:12px;font-weight:600;padding:2px 9px;border-radius:2px;white-space:nowrap}\n.b-green{background:var(--green-bg);color:var(--green)} .b-amber{background:var(--amber-bg);color:var(--amber)} .b-grey{background:var(--grey);color:var(--mid)}\n.ess{display:inline-block;font-size:12px;font-weight:700;padding:2px 9px;border-radius:2px;white-space:nowrap}\n.ess-core{background:var(--green);color:var(--surface)} .ess-opt{background:var(--green-bg);color:var(--green);box-shadow:inset 0 0 0 1px var(--green-line)}\n.cond{font-family:var(--mono);font-size:12px;color:var(--mid);margin-top:4px}\n.flag{color:var(--amber);font-weight:600}\n.legend{display:flex;flex-wrap:wrap;gap:8px 20px;font-size:13px;color:var(--mid);margin-top:10px}\n.legend span{display:inline-flex;align-items:center;gap:6px}\n.feature{background:var(--surface);border:1px solid var(--rule);border-radius:12px;margin-bottom:28px;overflow:hidden;scroll-margin-top:12px}\n.f-head{padding:20px 24px;border-bottom:1px solid var(--rule);display:grid;grid-template-columns:auto 1fr;gap:6px 16px}\n.fnum{grid-row:span 2;display:grid;place-items:center;width:34px;height:34px;border-radius:50%;background:var(--brand);color:#fff;font-weight:700}\n.ftext{margin:0;font-size:17px;line-height:1.75;color:var(--ink)}\n.fverdict{display:flex;flex-wrap:wrap;gap:8px;align-items:center}\n.cond-pill{font-family:var(--mono);font-size:12px;color:var(--mid);border:1px solid var(--rule);padding:1px 8px;border-radius:2px}\n.side{font-size:12.5px;color:var(--mid)}\nmark.hl{color:inherit;border-radius:3px;padding:1px 3px;box-decoration-break:clone;-webkit-box-decoration-break:clone}\nmark.e1{background:var(--e1);box-shadow:inset 0 -3px 0 var(--e1l)} mark.e2{background:var(--e2);box-shadow:inset 0 -3px 0 var(--e2l)} mark.e3{background:var(--e3);box-shadow:inset 0 -3px 0 var(--e3l)}\nmark.hl sup{font:700 10px/1 var(--sans);margin-left:2px}\nmark.e4{background:var(--e4);box-shadow:inset 0 -3px 0 var(--e4l)} mark.e5{background:var(--e5);box-shadow:inset 0 -3px 0 var(--e5l)}\nmark.e4 sup{color:var(--e4l)} mark.e5 sup{color:var(--e5l)}\nmark.e1 sup{color:var(--e1l)} mark.e2 sup{color:var(--e2l)} mark.e3 sup{color:var(--e3l)}\n.elements{margin:0} .elements caption{text-align:left;padding:16px 24px 6px;font-weight:600;color:var(--heading)}\n.elements th,.elements td{text-align:left;vertical-align:top;padding:10px 12px;border-top:1px solid var(--rule);font-size:14px}\n.elements thead th{font-size:12.5px;font-weight:600;color:var(--muted);border-top:0}\n.elements td:first-child,.elements th:first-child{padding-left:24px;width:44px} .elements td:last-child{padding-right:24px}\n.el{font-weight:600} .st{white-space:nowrap} .exl a{display:block;font-size:13px;white-space:nowrap}\n.sw{display:inline-grid;place-items:center;width:22px;height:22px;border-radius:4px;font-size:11px;font-weight:700}\n.sw.e1{background:var(--e1);color:var(--e1l);box-shadow:inset 0 0 0 1px var(--e1l)} .sw.e2{background:var(--e2);color:var(--e2l);box-shadow:inset 0 0 0 1px var(--e2l)} .sw.e3{background:var(--e3);color:var(--e3l);box-shadow:inset 0 0 0 1px var(--e3l)}\n.sw.e4{background:var(--e4);color:var(--e4l);box-shadow:inset 0 0 0 1px var(--e4l)} .sw.e5{background:var(--e5);color:var(--e5l);box-shadow:inset 0 0 0 1px var(--e5l)}\n.sw.sw0{background:var(--grey)}\n.caveat{margin:4px 24px 0;background:var(--amber-bg);border:1px solid var(--amber-line);border-radius:6px;padding:12px 16px}\n.cav-h{font-weight:700;color:var(--amber);font-size:13px} .caveat p{margin:4px 0 0;font-size:14px}\n.essline{margin:14px 24px 0;font-size:14px;color:var(--mid)} .essline strong{color:var(--heading)}\ndetails.more{border-top:1px solid var(--rule);margin-top:16px}\ndetails.more summary{cursor:pointer;padding:14px 24px;font-weight:600;color:var(--heading)}\ndetails.more summary .count{font-weight:400;color:var(--muted);font-size:13px;margin-left:8px}\ndetails.more>h4{margin:4px 24px 4px;font-size:13px;color:var(--muted)} details.more>p{margin:0 24px 14px;font-size:14px;color:var(--mid);max-width:80ch}\n.excerpt{margin:0 24px 16px;border:1px solid var(--rule);border-radius:6px;overflow:hidden;scroll-margin-top:12px}\n.excerpt figcaption{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;background:var(--surface-alt);padding:8px 14px;font-size:12.5px;font-weight:600;color:var(--heading)}\n.excerpt .ref{font-family:var(--mono);font-weight:400;color:var(--muted)}\n.exh{padding:10px 14px 0;font-weight:600;font-size:13.5px;color:var(--heading)}\n.excerpt pre{background:var(--excerpt-bg,var(--surface));margin:0;padding:8px 14px 14px;white-space:pre-wrap;overflow-wrap:anywhere;font:400 12.5px/1.65 var(--mono);color:var(--ink)}\n.appendix details{background:var(--surface);border:1px solid var(--rule);border-radius:6px;margin-bottom:12px}\n.appendix summary{cursor:pointer;padding:14px 20px;font-weight:600;color:var(--heading)}\n.appendix .inner{padding:0 20px 16px;font-size:14px;color:var(--mid)} .appendix dt{font-weight:600;color:var(--ink);margin-top:10px} .appendix dd{margin:2px 0 0}\n.intro{margin:-6px 0 18px;color:var(--mid);font-size:14px;max-width:80ch}\n.box .sub{display:block;font-size:12.5px;color:var(--muted);margin-bottom:4px}\na.fl{display:flex;align-items:center;gap:8px;text-decoration:none;padding:3px 0;color:var(--ink)} a.fl:hover{text-decoration:underline}\na.fl.in{display:inline-flex;font-size:13px;color:var(--mid);text-decoration:underline;text-underline-offset:2px;padding:0;margin-left:4px;white-space:nowrap}\n.fn{display:inline-grid;place-items:center;flex:none;width:20px;height:20px;border-radius:50%;background:var(--header);color:#fff;font-size:11px;font-weight:700}\n.vh{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0)}\nfooter{max-width:1040px;margin:0 auto;padding:16px 32px 40px;color:var(--muted);font-size:12.5px;border-top:1px solid var(--rule)}\n@media (max-width:800px){main{padding:0 16px 48px}.hdr-in{padding:18px 16px}.boxes,.boxes.two{grid-template-columns:1fr}.verdict{grid-template-columns:1fr}.ftext{font-size:16px}}\n@media print{details{display:block} details>summary{list-style:none} \n.hdr .logo svg{height:30px;width:auto;display:block}\n.verdict.v-amber{background:var(--amber-bg);border-color:var(--amber-line)} .verdict.v-amber .v-label{color:var(--amber)}\n.verdict.v-red{background:var(--red-bg);border-color:var(--red-line)} .verdict.v-red .v-label{color:var(--red)}\n.verdict.v-grey{background:var(--surface-alt);border-color:var(--rule)} .verdict.v-grey .v-label{color:var(--heading)}\n.ess-impl{background:var(--amber-bg);color:var(--amber)} .ess-no{background:var(--red-bg);color:var(--red)} .ess-inh{background:var(--grey);color:var(--mid)}\n.b-red{background:var(--red-bg);color:var(--red)}\n.cav-h.red{color:var(--red);margin-top:8px}\n.pill.restr{border-color:var(--amber-line);background:var(--amber-bg);color:var(--amber);text-decoration:none}\n.restricted{margin:0 0 12px;background:var(--amber-bg);border:1px solid var(--amber-line);border-left:3px solid var(--brand);border-radius:0 6px 6px 0;padding:12px 18px}\n.restricted .r-h{font-weight:700;color:var(--brand);font-size:12px;letter-spacing:.08em;text-transform:uppercase}\n.restricted p{margin:6px 0 0;color:var(--amber);font-size:13.5px}\nol.disc{margin:0;padding-left:20px} ol.disc li{margin:0 0 8px;font-size:13px}\n.exl a{display:block}\n@media print{.hdr,.brand-rule{-webkit-print-color-adjust:exact;print-color-adjust:exact} mark.hl,.sw,.badge,.ess,.verdict{-webkit-print-color-adjust:exact;print-color-adjust:exact}}\n";

// ── Helpers ────────────────────────────────────────────────────────────────
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const arr = (x) => Array.isArray(x) ? x : (x == null ? [] : [x]);

// Finds a quote inside text, ignoring case, spacing, punctuation and LaTeX
// markup. Returns [start, end) in the original text, or null.
function locate(text, quote) {
  const sig = [], map = [];
  const t = String(text ?? '');
  for (let i = 0; i < t.length; ) {
    if (t[i] === '\\' && /[A-Za-z]/.test(t[i + 1] || '')) { i++; while (i < t.length && /[A-Za-z]/.test(t[i])) i++; continue; }
    if (/[A-Za-z0-9]/.test(t[i])) { sig.push(t[i].toLowerCase()); map.push(i); }
    i++;
  }
  const q = String(quote ?? '').replace(/\\[A-Za-z]+/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (q.length < 3) return null;
  const k = sig.join('').indexOf(q);
  if (k < 0) return null;
  return [map[k], map[k + q.length - 1] + 1];
}

// Wraps ranges [{s, e, cls, n}] in <mark>, escaping everything else.
function markRanges(text, ranges, withSup) {
  const t = String(text ?? '');
  const rs = ranges.filter(Boolean).sort((a, b) => a.s - b.s);
  let out = '', pos = 0;
  for (const r of rs) {
    if (r.s < pos) continue;                      // overlapping: keep the first
    out += esc(t.slice(pos, r.s)) + `<mark class="hl ${r.cls}">` + esc(t.slice(r.s, r.e)) + (withSup ? `<sup>${r.n}</sup>` : '') + '</mark>';
    pos = r.e;
  }
  return out + esc(t.slice(pos));
}

function parseExcerptV2(str) {
  const s = String(str ?? '');
  const num = ((s.match(/\*\*Excerpt_Number:\*\*\s*([^\s]+)/) || [])[1] || '?').replace(/\.$/, '');
  const m = s.match(/\*\*Excerpt_Text:\*\*\s*Excerpt:[ \t]*\n([\s\S]+)/);
  let body = (m ? m[1] : s).replace(/\n---[ \t]*$/, '').trim();
  const ref = ((body.match(/Reference:[ \t]*\n\*\*([^*\n]+)\*\*/) || body.match(/Reference:[ \t]*\n([^\n]+)/) || [])[1] || '').trim();
  body = body.replace(/\n?Reference:[ \t]*\n[^\n]*$/, '').trim();
  const lines = body.split('\n');
  let heading = '';
  const kept = [];
  for (const l of lines) {
    if (/^#\s/.test(l.trim())) continue;                                   // running-header artefact
    if (!heading && /^##\s+/.test(l.trim())) { heading = l.trim().replace(/^##\s+/, ''); continue; }
    kept.push(l);
  }
  return { num, ref, heading, text: kept.join('\n').replace(/\n{3,}/g, '\n\n').trim() };
}

const shortDisclosure = (d) => ({ 'Explicitly Disclosed': 'Explicit', 'Implicitly Disclosed': 'Implicit', 'Functionally Equivalent': 'Equivalent', 'Not Disclosed': 'Not disclosed' }[d] || d || '');
function essParts(cls) {
  const [scope, side] = String(cls || '').split(' | ');
  const m = scope.match(/^Essential \((.+)\)$/);
  const label = m ? m[1].replace(/ in Main Profile$/, '').replace(/^Optional in Profile SEP$/, 'Optional (profile SEP)') : scope;
  const kind = m ? (/^Core/.test(m[1]) ? 'core' : 'opt') : (/inherent|non-technical/i.test(scope) ? 'inh' : /implementation/i.test(scope) ? 'impl' : 'no');
  return { scope, side: side || '', label, kind };
}
const sideWords = (side) => ({ 'Encoder and Decoder': 'Encoder and decoder', 'Decoder': 'Decoder only', 'Encoder': 'Encoder only', 'Codec System': 'Codec system only' }[side] || '');
function featList(xs) {
  const a = arr(xs).map(String);
  if (!a.length) return '';
  if (a.length === 1) return `feature ${a[0]}`;
  return `features ${a.slice(0, -1).join(', ')} and ${a[a.length - 1]}`;
}
const cap = (s) => s ? s[0].toUpperCase() + s.slice(1) : s;
const flink = (n) => `<a class="chip" href="#f${esc(n)}">Feature ${esc(n)}</a>`;
const STATUS_BADGE = { 'Explicit': 'b-green', 'Implicit': 'b-green', 'Equivalent': 'b-amber', 'Not found': 'b-red', 'Inherent': 'b-grey', 'Framing': 'b-grey' };
const STATUS_LABEL = { 'Not found': 'Not disclosed', 'Framing': 'Claim wording' };
// Summary-table caveat: the feature's gap, or its disclosure when that is itself the caveat
function caveatOf(gap, disclosure) {
  if (gap && gap !== 'None') return gap;
  if (disclosure === 'Not Disclosed') return 'Not disclosed in the standard';
  if (disclosure === 'Functionally Equivalent') return 'Relies on functional equivalence';
  return '';
}

// ── Builder ────────────────────────────────────────────────────────────────
function buildHtmlV2(data, meta, restricted) {
  const snap    = data.Snapshot || {};
  const verdict = snap.verdict || {};
  const align   = snap.alignment || {};
  const charts  = arr(data.Claim_Charts);
  const sumByIdx = {};
  for (const f of arr(snap.features)) sumByIdx[String(f.index)] = f;

  const patent   = data.Patent_Number || meta.Patent_Number || '';
  const title    = data.Title || meta.Title || 'Patent analysis report';
  const owner    = data.Owner || meta.Owner || '';
  const standard = [data.Standard || meta.Standard || '', data.Standard_Edition || ''].filter(Boolean).join(' · ');
  const claimNo  = data.Claim_Number || '';
  const catWords = { Encoder: 'Encoder claim', Decoder: 'Decoder claim', Encoder_and_Decoder_System: 'Encoder and decoder claim', Other: 'Other claim type' }[data.Claim_Category] || '';

  // verdict
  const ov = essParts((verdict.classification || '') + (verdict.side ? ' | ' + verdict.side : ''));
  const vClass = ov.kind === 'core' || ov.kind === 'opt' ? 'v-green' : ov.kind === 'impl' ? 'v-amber' : ov.kind === 'inh' ? 'v-grey' : 'v-red';
  const vLabel = (ov.kind === 'core' || ov.kind === 'opt') ? `Essential · ${esc(ov.scope.replace(/^Essential \((.+)\)$/, '$1'))}` : esc(ov.scope || '—');
  const bottom = verdict.bottom_line || '';

  // where it aligns
  const parts = [];
  if (arr(align.stated_directly).length)  parts.push(`${cap(featList(align.stated_directly))} ${align.stated_directly.length > 1 ? 'are' : 'is'} stated directly in the standard`);
  if (arr(align.by_implication).length)   parts.push(`${featList(align.by_implication)} ${align.by_implication.length > 1 ? 'follow' : 'follows'} by necessary implication`);
  if (arr(align.by_equivalence).length)   parts.push(`${featList(align.by_equivalence)} ${align.by_equivalence.length > 1 ? 'map' : 'maps'} by functional equivalence`);
  if (arr(align.not_disclosed).length)    parts.push(`${featList(align.not_disclosed)} ${align.not_disclosed.length > 1 ? 'are' : 'is'} not disclosed`);
  if (arr(align.inherent).length)         parts.push(`${featList(align.inherent)} ${align.inherent.length > 1 ? 'are' : 'is an'} inherent component${align.inherent.length > 1 ? 's' : ''}`);
  const lead = parts.length ? cap(parts.join('; ')) + '.' : '';
  const provRows = arr(align.provisions).map(p => `<li><b>${esc(p.clause)}</b><span class="pd">${esc(p.title)}${p.verify ? ' <span class="tag ver">Verify</span>' : ''}</span><span class="chips">${arr(p.features).map(flink).join('')}</span></li>`).join('');

  // gaps and conditions
  const G = arr(snap.gaps_and_conditions);
  const tagCls = { Gap: 'gap', Equivalence: 'ver', Condition: 'cond', Verify: 'ver', 'Non-required': 'gap', Product: 'ver' };
  const gapRows = [];
  if (!G.some(g => g.type === 'Gap')) gapRows.push(`<li><span class="tag gap">Gap</span>None</li>`);
  for (const g of G) {
    let txt = esc(g.text);
    if (g.type === 'Condition' && arr(g.flags).length) txt = 'Enabling flag: ' + arr(g.flags).map(f => `<code>${esc(f)}</code>`).join(' / ');
    gapRows.push(`<li><span class="tag ${tagCls[g.type] || 'ver'}">${esc(g.type)}</span>${txt}${g.feature ? ' ' + flink(g.feature) : ''}</li>`);
  }

  // summary mapping
  const sumRows = charts.map(c => {
    const i = String(c?.Claim_Feature?.Index ?? '');
    const s = sumByIdx[i] || {};
    const dec = c.Decision || {};
    const e = essParts(dec.Essentiality_Classification);
    const flags = arr(dec.Gating_Flags);
    const cav = caveatOf(s.gap, dec.Disclosure);
    const gap = cav ? `<span class="flag">${esc(cav)}</span>` : '<span class="none">None</span>';
    return `<tr><th scope="row"><a href="#f${esc(i)}">${esc(i)}</a></th><td>${esc(s.short_label || c?.Claim_Feature?.Text || '')}</td><td class="maps">${esc(s.maps_to || '')}</td>` +
      `<td><span class="ess ess-${e.kind}">${esc(e.label)}</span>${flags.length ? `<div class="cond">${esc(flags.join(' / '))}</div>` : ''}</td><td>${gap}</td></tr>`;
  }).join('');

  // detailed chart
  const cards = charts.map(c => {
    const i = String(c?.Claim_Feature?.Index ?? '');
    const s = sumByIdx[i] || {};
    const dec = c.Decision || {};
    const an  = c.Analysis || {};
    const e   = essParts(dec.Essentiality_Classification);
    const els = arr(dec.Elements);
    const ftext = c?.Claim_Feature?.Text || '';
    const ex = arr(c.Cited_Excerpts).map(parseExcerptV2);
    const exByNum = {}; ex.forEach(x => exByNum[x.num] = x);

    // colours for non-framing elements
    let ci = 0;
    const colored = els.map(el => ({ el, n: el.status === 'Framing' ? 0 : (++ci), cls: el.status === 'Framing' ? '' : `e${((ci - 1) % 5) + 1}` }));
    const ftRanges = colored.filter(x => x.n).map(x => {
      let r = locate(ftext, x.el.element_text) || locate(ftext, String(x.el.element_text || '').replace(/^\s*\d+\.\s*/, ''));
      return r ? { s: r[0], e: r[1], cls: x.cls, n: x.n } : null;
    });
    const exRanges = {};   // num -> ranges
    const citedNums = new Set();
    for (const x of colored) for (const ev of arr(x.el.evidence)) {
      const num = String(ev.excerpt_number); citedNums.add(num);
      const tx = exByNum[num]; if (!tx || !x.n) continue;
      const r = locate(tx.text, ev.quote);
      if (r) (exRanges[num] = exRanges[num] || []).push({ s: r[0], e: r[1], cls: x.cls, n: x.n });
    }
    const eRows = colored.map(x => {
      const sw = x.n ? `<span class="sw ${x.cls}">${x.n}</span>` : '<span class="sw sw0"></span>';
      const links = [...new Set(arr(x.el.evidence).map(ev => String(ev.excerpt_number)))].map(n => `<a href="#f${esc(i)}-x${esc(n)}">Excerpt ${esc(n)}</a>`).join('') || '—';
      return `<tr><td>${sw}</td><td class="el">${esc(x.el.element_text)}</td><td>${esc(x.el.maps_to || '')}</td>` +
        `<td class="st"><span class="badge ${STATUS_BADGE[x.el.status] || 'b-grey'}">${esc(STATUS_LABEL[x.el.status] || x.el.status)}</span>${x.el.verify ? ' <span class="badge b-amber">Verify</span>' : ''}</td><td class="exl">${links}</td></tr>`;
    }).join('');
    const verify = arr(dec.Verify);
    const notFound = colored.filter(x => x.el.status === 'Not found');
    const caveat = (verify.length || notFound.length) ? `<div class="caveat">${verify.length ? `<div class="cav-h">To verify</div>${verify.map(v => `<p>${esc(v)}</p>`).join('')}` : ''}${notFound.length ? `<div class="cav-h red">Not found in the standard</div>${notFound.map(x => `<p>${esc(x.el.element_text)}</p>`).join('')}` : ''}</div>` : '';
    const essLine = s.essentiality_line || '';
    const excerptHtml = (x) => `<figure class="excerpt" id="f${esc(i)}-x${esc(x.num)}"><figcaption><span>Excerpt ${esc(x.num)}</span><span class="ref">${esc(x.ref)}</span></figcaption>${x.heading ? `<div class="exh">${esc(x.heading)}</div>` : ''}<pre>${markRanges(x.text, exRanges[x.num] || [], true)}</pre></figure>`;
    const cited = ex.filter(x => citedNums.has(x.num)), other = ex.filter(x => !citedNums.has(x.num));
    const reasoning = [['Interpretation', an.Interpretation], ['Opinion', an.Overall_Opinion], ['Differences', an.Differences && !/^none\.?$/i.test(String(an.Differences).trim()) ? an.Differences : ''], ['Essentiality justification', dec.Justification]]
      .filter(([, v]) => v).map(([h, v]) => `<h4>${h}</h4><p>${esc(v)}</p>`).join('');
    const dcls = dec.Disclosure === 'Not Disclosed' ? 'b-red' : dec.Disclosure === 'Functionally Equivalent' ? 'b-amber' : 'b-green';
    return `<article class="feature" id="f${esc(i)}">
 <header class="f-head"><span class="fnum">${esc(i)}</span><p class="ftext">${markRanges(ftext, ftRanges, true)}</p>
  <div class="fverdict"><span class="badge ${dcls}">${esc(dec.Disclosure || '')}</span><span class="ess ess-${e.kind}">${esc(e.label)}</span>${arr(dec.Gating_Flags).length ? `<span class="cond-pill">${esc(arr(dec.Gating_Flags).join(' / '))}</span>` : ''}${sideWords(e.side) ? `<span class="side">${sideWords(e.side)}</span>` : ''}</div></header>
 ${els.length ? `<table class="elements"><caption>How each element maps</caption><thead><tr><th scope="col"><span class="vh">Element</span></th><th scope="col">Claim element</th><th scope="col">Maps to</th><th scope="col">Disclosure</th><th scope="col">Evidence</th></tr></thead><tbody>${eRows}</tbody></table>` : ''}
 ${caveat}
 ${essLine ? `<p class="essline"><strong>Essentiality.</strong> ${esc(essLine)}</p>` : ''}
 ${reasoning ? `<details class="more"><summary>Reasoning</summary>${reasoning}</details>` : ''}
 ${cited.length ? `<details class="more" open><summary>Cited excerpts <span class="count">${cited.length}</span></summary>${cited.map(excerptHtml).join('')}</details>` : ''}
 ${other.length ? `<details class="more"><summary>Other retrieved excerpts <span class="count">${other.length}</span></summary>${other.map(excerptHtml).join('')}</details>` : ''}
</article>`;
  }).join('\n');

  // appendix
  const lim = String(data['Limitation(s)'] || '');
  const limLabel = lim.split('\n')[0].trim();
  const limBody = lim.split('\n').slice(1).join('\n').trim();
  const M = data.Methodology || {};
  const terms = [...arr(M.disclosure_categories), ...arr(M.essentiality_tiers)].map(t => `<dt>${esc(t.label)}</dt><dd>${esc(t.definition)}</dd>`).join('') +
    Object.entries(M.universal_metrics || {}).map(([k, v]) => `<dt>${esc({ percentage_mapped: 'Mapped', weighted_mapping: 'Evidence strength (weighted mapping)' }[k] || k)}</dt><dd>${esc(v)}</dd>`).join('');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(patent)} · Claim ${esc(claimNo)} · Claim chart</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=Source+Sans+3:wght@400;500;600;700&family=Source+Code+Pro:wght@400;500&display=swap" rel="stylesheet">
<style>${CSS_V2}</style>
</head>
<body>
<div class="brand-rule"></div>
<div class="hdr"><div class="hdr-in"><div class="logo" aria-label="IP Mind">${LOGO_SVG}</div><div class="conf">Confidential</div></div></div>
<main>
<section class="ident">
 <div class="pills">${patent ? `<span class="pill">${esc(patent)}</span>` : ''}${claimNo ? `<span class="pill">Claim ${esc(claimNo)}</span>` : ''}${standard ? `<span class="pill std">${esc(standard)}</span>` : ''}${restricted ? `<a class="pill restr" href="#restricted">Restricted use</a>` : ''}</div>
 <h1>${esc(title)}</h1>
 <p class="owner">${[owner, catWords].filter(Boolean).map(esc).join(' · ')}</p>
 ${data.Claim ? `<div class="claim"><div class="claim-h">Claim ${esc(claimNo)}</div><p>${esc(data.Claim)}</p></div>` : ''}
</section>

<h2>At a glance</h2>
<section class="verdict ${vClass}" aria-label="Verdict">
 <div><p class="v-label">${vLabel}</p>${bottom ? `<p class="v-line">${esc(bottom)}</p>` : ''}</div>
 <div class="v-stats"><div class="stat"><b>${esc(verdict.percentage_mapped ?? '—')}%</b><span>Mapped</span></div><div class="stat"><b>${esc(verdict.weighted_percentage_mapped ?? '—')}%</b><span>Evidence strength</span></div><div class="stat"><b>${esc(verdict.features_aligned ?? '—')}/${esc(verdict.features_technical ?? '—')}</b><span>Features aligned</span></div></div>
</section>
<div class="boxes two">
 <section class="box align"><h3>Where it aligns</h3>${lead ? `<p class="lead">${esc(lead)}</p>` : ''}<ul class="prov">${provRows}</ul></section>
 <section class="box gaps"><h3>Gaps and conditions</h3><ul>${gapRows.join('')}</ul></section>
</div>

<h2>Summary mapping</h2>
<p class="intro">One line per feature. The full analysis and evidence for each is in the detailed claim chart below.</p>
<div class="tablewrap"><table class="align-t"><thead><tr><th scope="col">#</th><th scope="col">Claim feature</th><th scope="col">Maps to</th><th scope="col">Essentiality</th><th scope="col">Gap or caveat</th></tr></thead><tbody>${sumRows}</tbody></table></div>
<div class="legend"><span><span class="ess ess-core">Core</span> required in every Main / Main 10 bitstream</span><span><span class="ess ess-opt">Optional</span> essential when an enabling flag is on</span><span><span class="badge b-amber">Verify</span> rests on a provision not in the cited excerpts</span></div>

<h2>Detailed claim chart</h2>
<p class="intro">Each element of a feature has its own colour and number. The same colour and number mark the supporting text in the excerpts.</p>
${cards}

<h2>Appendix</h2>
<section class="appendix">
 ${data.Summary ? `<details><summary>Full conclusion</summary><div class="inner"><p>${esc(data.Summary)}</p></div></details>` : ''}
 ${lim ? `<details><summary>Limitations${limLabel ? ': ' + esc(limLabel) : ''}</summary><div class="inner"><p>${esc(limBody || limLabel)}</p></div></details>` : ''}
 ${terms ? `<details><summary>Key to terms</summary><div class="inner"><dl>${terms}</dl></div></details>` : ''}
 ${restricted ? `<div class="restricted" id="restricted"><div class="r-h">Restricted use notice</div><p>${esc(RESTRICTED_NOTICE)}</p></div>` : ''}
 <details open><summary>Disclaimer</summary><div class="inner"><ol class="disc">${DISCLAIMER_ITEMS_HTML.map(x => `<li>${x}</li>`).join('')}</ol></div></details>
</section>
</main>
<footer>ipmind.ai</footer>
<script>window.addEventListener('beforeprint',function(){document.querySelectorAll('details').forEach(function(d){d.setAttribute('open','')})});</script>
</body></html>`;
}

module.exports = { buildHtmlV2, locate, parseExcerptV2 };
