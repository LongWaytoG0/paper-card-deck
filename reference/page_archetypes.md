# 页型档案 · Page Archetypes (quick reference)

A paper module = **simple (2–3p)** + **detailed (3–6p)**, same template chrome.
Both versions are assembled from the archetypes below. All coords in inches on a
13.333 × 7.5 slide. Colours/fonts: see `scripts/pptx_cards.py` (`C`, `CN`, `EN`, `FRAME`).

Golden rule: **reuse the paper's own figures/tables as cropped images** + Chinese
explanation & highlight callouts. Do NOT rebuild dense result tables natively.

---

## Shared chrome (every page)
- Running tag, top-left `y≈0.09`: `▪ CONF YEAR · <ShortTitle>：<中文副标>` (navy square marker + two-tone text, 9.5pt).
- Logos top-right: NWPU `x=9.101`, ASLP `x=10.555`, both `y=0.20`.
- Page number bottom-right `x=12.20 y=7.02`, faint.
- Footer line (author list + links) bottom-left `y≈6.96`, 10pt:
  `Authors. CONF YEAR Oral. Paper: <url> · Demo: <url> · GitHub: <url>`
  Show ONLY the links actually provided — omit any that weren't given, and never repurpose
  another URL (e.g. a leaderboard) as a demo link.
- Interior pages add a header: blue kicker `y=0.42` (e.g. `01 · BENCHMARK DESIGN`) + navy title `y=0.72` (25pt).

## 1 · Cover / Overview  (封面)
Two-column "背景 + 方案" layout (the canonical cover):
- Big two-tone title `y=0.52` (27pt) + English subtitle `y=1.22` (13.5pt italic mute) + hairline divider `y=1.66`.
- **Left `x=0.55 w=5.75`** — heading「研究背景与动机 · Background」; 3–4 bullets, each a colour square marker + 【小标题】(bold accent) + mute explanation. Last bullet = 研究动机.
- **Right `x=6.55 w=6.18`** — heading「提出方案 · <ShortName>」; the paper's framework figure on top (`figure()`), caption, then 3 contribution one-liners (colour dot + bold EN term + · CN gloss).
- Alt cover (used in detailed v1 earlier): single motivation paragraph full-width + 3 numbered contribution cards (left) + framework figure (right).

## 2 · Method / Design  (方法)
Left + right columns. Typical: left = a stacked list of taxonomy/model axes
(panel rows: bold EN key + mute value). Right = numbered pipeline steps
(`num_dot` + EN title + CN description) + a highlighted example box at the bottom.

## 3 · Protocol / Pipeline  (协议)
Full-width banner figure of the paper's pipeline (`figure()`, centered), then a row
of 3 stage cards below: accent dot + EN stage title + CN one-liner + a formula chip
(white rounded rect, accent border) holding the key equation.

## 4 · Results & Findings  (结果)
Left = the paper's main results **table image** (cropped PNG) in a white frame, with a
Chinese caption; optional translucent highlight rectangles over the key cells. Right =
3–4 finding cards (panel + accent dot + bold navy title + mute sentence). This is the
page that most relies on reusing the paper image rather than rebuilding.

## 5 · Audio Demo  (音频 Demo，占位)
2×2 cards. Each: coloured category chip (e.g. `Timbre · Age = Child`), difficulty tag
(right), `INSTRUCTION` label + quoted instruction, `SYNTHESIS TEXT` label + quoted text,
and a **dashed audio/QR placeholder box** (`▶ 音频 / 二维码占位 · 在此插入试听`). Fonts
match interior pages (chip 12.5pt, labels 10.5pt, quotes 13pt). Online-demo link at top-right.

## 6 · Video Demo  (视频 Demo，占位)
Left: large 16:9 **dashed video placeholder** with a navy play button + caption. Right:
「Highlights」list of 3–4 chapter cards (accent dot + bold title + CN one-liner). Online-demo link top-right.

---

## Simple vs. detailed — which archetypes
- **Simple (2–3p):** Cover + Results (+ optionally one Demo page). Pick the single most
  important figure and the headline findings.
- **Detailed (3–6p):** Cover + Method + Protocol + Results + Audio Demo + Video Demo (drop
  any that don't apply to the paper).

Page numbers may be inherited from the detailed deck when a simple deck is a subset of
slides (e.g. simple shows 1 / 4 / 6) — that is acceptable, leave as the user set it.
