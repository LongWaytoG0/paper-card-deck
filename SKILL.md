---
name: paper-card-deck
description: >-
  Turn a research paper into an ASLP@NPU lab-intro "paper-card" slide deck — a compact
  simple version (2–3 pages) AND a detailed version (3–6 pages) on one shared branded
  template. Use when the user gives a paper PDF (+ paper / demo / GitHub links) and wants
  lab-style introduction slides, paper cards, or a module to merge into a combined lab deck.
  Input = paper PDF + links; output = the simple and detailed .pptx files.
---

# Paper-Card Deck (ASLP@NPU lab-intro slides)

Build a lab-introduction **paper module**: each paper becomes a compact, self-contained
set of slides in **two versions** on the same template —
a **simple version (2–3 pages)** and a **detailed version (3–6 pages)**.
Modules from different papers must look identical (same header/footer) so they can be
**merged into one combined lab deck**: important papers contribute their detailed module,
minor papers their simple module.

> **This SKILL.md is self-contained.** The full helper library is embedded below (Appendix A).
> Do the **Bootstrap** step first. Real logo bitmaps ship in the `.skill` bundle; without them
> the header falls back to clean text brand-marks, so nothing breaks.

## Input → Output contract
**Inputs** (from the user):
- a **paper PDF** (required) — the source of all content, figures, and tables;
- a **paper link** (arXiv/DOI); optionally a **demo link** and a **GitHub link**.

**Outputs** (deliver both):
- `<ShortName>_详细版.pptx` — detailed module (3–6 pages);
- `<ShortName>_简版.pptx` — simple module (2–3 pages, a curated subset).

Ask only what you truly can't infer from the PDF (e.g. which figures matter most, or how
many pages).

**Proven default structure (use unless the user says otherwise):** detailed = **5 pages =
4 paper pages + 1 demo page**:
①封面（背景动机 + 方案图）②方法总览（核心方法 + 主图 + 公式/说明卡）③方法细节（按这篇论文灵活
定制的关键设计页）④结果 ⑤demo。The simple version is **not rebuilt** — it **reuses the exact same
page functions** as a 4-page subset: 封面 + 方法总览 + 结果 + demo (just renumber). Write each
page as a function `page(prs, pageno)` so both decks call the same code; the simple deck is
`cover(d); p_method(d,2); p_results(d,3); p_demo(d,4)`. Pages ② and ③ are the flexible slots —
tailor them to the paper (for an audio-codec paper, e.g., a dedicated "5 Hz VAE" design page).

## Bootstrap (do this FIRST, every run)
1. Save the **Appendix A** code block to `pptx_cards.py` in your working dir.
2. `from pptx_cards import *`, then call `header_logos(slide)` on every page to place the
   top-right marks.
3. **Logos:** the real ASLP@NPU + NWPU logo bitmaps ship inside the `.skill` bundle's
   `scripts/pptx_cards.py` (embedded as base64, written out by `write_logos('assets')`).
   The Appendix A copy below is the same library with the bitmaps omitted — if that is all
   you have, `header_logos()` automatically draws clean **text brand-marks** instead, so the
   header never breaks. To get the real logos: use the bundled `scripts/pptx_cards.py` if the
   skill directory has it, or ask the user for the logo PNGs / a template that contains them,
   drop them in `assets/logo_aslp.png` (NWPU mark) + `assets/logo_nwpu.png` (ASLP mark).


## Core conventions (do not violate)
1. **Reuse the paper's own figures & tables as cropped images**, then add Chinese
   explanation + highlight callouts. Do **NOT** rebuild dense result tables natively —
   crop the table from the PDF and overlay translucent highlight rectangles on key cells.
2. **Language style = 中文讲解 + 英文术语混排.** Narrate in Chinese; keep English for
   technical terms (taxonomy, flow matching, VAE, CPPO, WER, IF/PE score, etc.).
3. **2–6 pages per module, never a 15-page talk.** Each page must instantly convey: what
   the paper is, its contributions, one key figure, the headline result.
4. **Preserve the template chrome exactly** on every page: ASLP@NPU + NWPU logos (top-right),
   running tag (top-left), page number (bottom-right), footer author+links line (bottom-left).
5. **Never regenerate-and-overwrite a file the user may have hand-edited.** If iterating on a
   user-tuned deck, edit it **in place** (python-pptx) and back it up first; do not rebuild
   from scratch and clobber their manual font/spacing tuning.
6. **Only show links the user actually gave you — never fabricate or repurpose one.** Put each
   link under its true label (Paper / Demo / GitHub / Leaderboard / Project page). If no demo
   link was provided, show **no** demo link and add **no** "在线试听 / online-demo" pill — do
   not dress up a different URL (e.g. a leaderboard or project page) as a demo. An absent link
   is simply omitted; a placeholder audio/QR box on a demo page is fine (it's clearly a
   placeholder, not a claimed link). When unsure what a URL is, ask rather than guess its label.

## Workflow
1. **Read the paper** with the `pdf-reading` skill. Extract: title, author list, the core
   problem/motivation, the method/framework, ~3 contributions, the pipeline, the main
   results table(s), and the headline findings.
2. **Harvest figures/tables from the PDF** as PNGs (rasterize the relevant page region with
   `pdftoppm`/`pdftocairo`, then crop to the figure's bounding box; or extract embedded
   images). These become the figures on the Method / Protocol / Results pages.
3. **Pick the page set** for detailed and simple (see Page Archetypes below).
4. **Build on the template** using the embedded helper:
   - Start a 13.333×7.5 deck; for each page create a branded blank slide (logos + running
     tag + marker + page number + footer line), then lay out the chosen archetype.
   - When iterating on a user file instead, duplicate a branded slide with the pptx skill's
     `scripts/add_slide.py`, then `strip_to_template()` it and add content (preserves tuning).
5. **Footer links line** on every page, showing **only the links actually provided**:
   `Authors. <CONF YEAR> Oral. Paper: <url> · Demo: <url> · GitHub: <url>`
   (omit any not given — do not invent or substitute). Only add an online-demo pill top-right
   on a demo page **if a real demo link exists**; otherwise leave the demo page link-free.
6. **Validate + visual QA** (below). Deliver both `.pptx` via `computer://` links.
7. **Offer** to merge the module into the user's combined lab deck.

## Page archetypes
A paper module is assembled from these. All coords in inches on a 13.333×7.5 slide.
Reuse the paper's figures/tables as cropped PNGs + Chinese explanation & highlight callouts.

- **Cover (封面):** two-column "背景 + 方案". Big two-tone title `y=0.52` (27pt) + English
  subtitle `y=1.22` (13.5pt italic mute) + hairline divider `y=1.66`. Left `x=0.55 w=5.75`
  = 「研究背景与动机 · Background」+ 3–4 bullets (colour square marker + 【小标题】bold accent
  + mute text; last = 研究动机). Right `x=6.55 w=6.18` = 「提出方案 · <Name>」+ framework
  figure on top (`figure()`) + caption + 3 contribution one-liners (dot + bold EN + · CN).
- **Method overview (方法总览, page ②):** left column = the paper's core-method figure
  (`place_fig`, ~`x=0.5 w=7.1`) + a formula chip under it (white rounded rect, accent border,
  centered equation). Right column = 4 explanation cards (`rx≈7.9 rw≈4.9`, panel + accent
  square + bold navy title + mute sentence, `ch≈1.2 gap≈0.15`) each naming one mechanism.
- **Design deep-dive (方法细节 / 关键设计, page ③ — tailor to THIS paper):** a flexible slot.
  Pick the one design that most defines the paper and build a dedicated page: one intro line,
  left = the relevant **table/figure crop** (often with a highlighted row, see `hi_row`) plus a
  row of **key-number chips** (`big` stat + small caption), right = 2–3 design cards. (For
  DiffRhythm 2 this was the "5 Hz Music VAE" reconstruction-comparison page.)
- **Method / Design (generic 方法):** left+right columns. Left = stacked list of model/taxonomy axes
  (panel rows: bold EN key + mute value). Right = numbered pipeline steps (`num_dot` + EN
  title + CN desc) + a highlighted example box at the bottom.
- **Protocol / Pipeline (协议):** full-width banner figure of the pipeline (`figure()`,
  centered), then a row of 3 stage cards: accent dot + EN stage title + CN one-liner +
  a formula chip (white rounded rect, accent border) holding the key equation.
- **Results & Findings (结果):** left = the paper's main results **table image** (cropped PNG)
  in a white frame + Chinese caption + translucent highlight **bands over this paper's own
  row(s)** so the reader instantly sees where it wins (use `hi_row(slide, img, x, y, w,
  center_frac, col)` — pass the row's vertical center as a fraction of the image height;
  measure it from the LEFT label column and watch for sub-header rows). Right = 3–4 finding
  cards (panel + accent dot + bold navy title + mute sentence).
- **Audio Demo (音频 Demo, 占位):** 2×2 cards. Each: coloured category chip, difficulty tag
  (right), `INSTRUCTION` label + quoted instruction, `SYNTHESIS TEXT` label + quoted text,
  and a **dashed audio/QR placeholder box**. Fonts: chip 12.5pt, labels 10.5pt, quotes 13pt.
  Online-demo pill top-right **only if a real demo link was provided** (else omit it).
- **Video Demo (视频 Demo, 占位):** left = large 16:9 **dashed video placeholder** + navy play
  button + caption. Right = 「Highlights」list of 3–4 chapter cards. Online-demo pill top-right
  **only if a real demo link was provided** (else omit it).

Accent-colour semantics (reuse consistently, keys in `C`): stage-1/taxonomy = `tax` (blue),
stage-2/data = `data` (teal), stage-3/eval = `eval` (amber), positive = `good`, bottleneck = `warn`.

**Which archetypes:** the proven default — detailed (5p) = cover + method-overview (②) +
design-deep-dive (③, tailored to the paper) + results + demo; simple (4p) = the SAME page
functions as a subset = cover + method-overview + results + demo (just renumber, don't
rebuild). Scale down to 2–3p for minor papers (cover + results [+ one demo]). Page numbers may
be inherited when the simple deck is a subset — that's fine.

## Validation & visual QA (always do this)
- **Validate** the saved pptx with the pptx skill's validator (pass `--original <template>`
  if you edited a template-derived deck in place).
- **Render & eyeball every page:** `soffice --headless --convert-to pdf --outdir . out.pptx`
  then `pdftoppm -jpeg -r 110 out.pdf page` and read the JPEGs.
- **Font caveat:** the Linux renderer usually lacks 微软雅黑, so LibreOffice substitutes a
  wider face — big CJK titles may wrap/overlap in the *preview* but are fine in real
  PowerPoint. Don't "fix" a wrap that's only a preview artifact; confirm before shrinking
  a title the user set.
- Previously-rendered PDFs/JPEGs can become lock-owned ("Operation not permitted" on rm).
  Render to a **fresh output filename** instead of deleting.

---

## Appendix A — helper library (`pptx_cards.py`)
Save this entire block to `pptx_cards.py`, then `from pptx_cards import *`. This is the
base64-free copy (logos omitted → `header_logos()` uses text brand-marks). The `.skill`
bundle ships the same file **with** the real logo bitmaps embedded.

```python
#!/usr/bin/env python3
"""
pptx_cards.py — reusable helper library for building ASLP@NPU-style
"paper-card" decks with python-pptx (edit-in-place on a template slide).

Design notes
------------
* Slide size is 13.333 x 7.5 in (16:9).  All coordinates below assume that.
* CJK fonts: python-pptx's run.font.name only sets <a:latin>. For 微软雅黑 (or any
  CJK face) to actually render you MUST also set <a:ea> and <a:cs>. set_font() does this.
* Default autoshapes carry a drop shadow — kill_shadow() removes it.
* Dashed borders / rounded-corner radius are set via raw XML (rrect()).

Typical usage
-------------
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx_cards import *          # C, CN, EN, set_font, add_text, rrect, ...

    prs = Presentation("template_with_one_slide.pptx")
    s = prs.slides[0]
    strip_to_template(s, "1")         # keep logos/tag/footer/pagenum, clear body
    add_text(s, 0.55, 0.72, 10, 0.6, [("标题 Title", dict(size=25, color=C["navy"], bold=True))])
    prs.save("out.pptx")
"""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette & fonts
# Brand blues sampled from the NWPU / ASLP logos, plus a semantic accent set.
C = {
    "navy":  "004479",  # NWPU deep blue — primary
    "navy2": "0A5A93",
    "blue":  "1E73C8",  # accent blue (kickers, links)
    "ink":   "1B2A38",  # near-black body
    "mute":  "5C6B7A",  # muted gray body
    "faint": "8A97A4",  # faint labels / page numbers
    "line":  "D5DEE8",  # hairlines / panel borders
    "panel": "F3F7FB",  # light panel fill
    "panel2":"E9F1F9",  # slightly deeper panel fill
    "white": "FFFFFF",
    "tax":   "1E73C8",  # taxonomy / stage-1 accent (blue)
    "data":  "0E9C8A",  # data-construction / stage-2 accent (teal)
    "eval":  "E08A2B",  # evaluation / stage-3 accent (amber)
    "good":  "1F9D57",  # positive highlight (green)
    "warn":  "D2483F",  # bottleneck / negative highlight (red)
    "tagmute":"5B7590",  # running-tag muted text
}
CN = "微软雅黑"   # Chinese face (also set as a:ea / a:cs)
EN = "Calibri"    # Latin face for English terms / numbers

# ---------------------------------------------------------------- template chrome
# Baseline coordinates for the shared header/footer. Tune per real template.
FRAME = {
    "logo_nwpu": dict(x=9.101, y=0.20, w=1.386, h=0.431),
    "logo_aslp": dict(x=10.555, y=0.20, w=2.168, h=0.43),
    "pagenum":   dict(x=12.20, y=7.02, w=0.90, h=0.30),   # bottom-right
    "tag":       dict(x=0.72, y=0.09, w=8.0, h=0.26),      # top-left running tag
    "tag_marker":dict(x=0.55, y=0.15, w=0.11, h=0.11),     # navy square before tag
    # interior page header
    "kicker":    dict(x=0.55, y=0.42, sz=11.5),            # e.g. "01 · BENCHMARK DESIGN"
    "title":     dict(x=0.55, y=0.72, sz=25),             # interior page title (navy)
    # cover page
    "cover_title":dict(x=0.55, y=0.52, sz=27),
    "subtitle":  dict(x=0.55, y=1.22, sz=13.5),
    "divider_y": 1.66,
    "margin": 0.55,
}

def rgb(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

# ---------------------------------------------------------------- run / text helpers
def set_font(run, size=None, color=None, bold=None, italic=None, name=CN):
    """Set font props AND patch a:latin/a:ea/a:cs so CJK faces render."""
    f = run.font
    if size is not None:  f.size = Pt(size)
    if bold is not None:  f.bold = bold
    if italic is not None: f.italic = italic
    if color is not None: f.color.rgb = rgb(color)
    if name is not None:
        f.name = name
        rPr = run._r.get_or_add_rPr()
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = rPr.makeelement(qn(tag), {}); rPr.append(el)
            el.set("typeface", name)

def kill_shadow(sh):
    try: sh.shadow.inherit = False
    except Exception: pass

def no_autofit_margins(tf):
    tf.word_wrap = True
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)

def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=None):
    """Add a textbox. `runs` = list of (text, opts). opts keys: size,color,bold,
    italic,name, and optional `link` (hyperlink URL). Use ("\\n",{}) for a new para."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; no_autofit_margins(tf); tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align
    if line_spacing: p.line_spacing = line_spacing
    for txt, opts in runs:
        if txt == "\n":
            p = tf.add_paragraph(); p.alignment = align
            if line_spacing: p.line_spacing = line_spacing
            continue
        opts = dict(opts); link = opts.pop("link", None)
        r = p.add_run(); r.text = txt; set_font(r, **opts)
        if link: r.hyperlink.address = link
    return tb

# ---------------------------------------------------------------- shape helpers
def rrect(slide, x, y, w, h, fill=None, line=None, lw=1.0, radius=0.08, dash=None):
    """Rounded rectangle. fill/line are hex strings or None(=transparent).
    dash e.g. 'dash' for dashed border. radius 0..0.5."""
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    kill_shadow(sp)
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = rgb(line); sp.line.width = Pt(lw)
    try: sp.adjustments[0] = radius
    except Exception: pass
    if dash is not None:
        ln = sp.line._get_or_add_ln()
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))
    sp.text_frame.word_wrap = True
    return sp

def dot(slide, x, y, d, color):
    """Small filled circle marker."""
    e = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    kill_shadow(e); e.fill.solid(); e.fill.fore_color.rgb = rgb(color); e.line.fill.background()
    return e

def num_dot(slide, x, y, n, color, d=0.44):
    """Numbered circle motif (e.g. contribution/step number)."""
    e = dot(slide, x, y, d, color)
    tf = e.text_frame; no_autofit_margins(tf); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = str(n); set_font(r, size=16, color=C["white"], bold=True, name=EN)
    return e

def figure(slide, img_path, x, y, w, h, caption=None, cap_sz=10.5):
    """Framed image with optional italic caption underneath."""
    rrect(slide, x-0.12, y-0.12, w+0.24, h+0.24+(0.28 if caption else 0),
          fill=C["white"], line=C["line"], lw=1, radius=0.03)
    slide.shapes.add_picture(img_path, Inches(x), Inches(y), Inches(w), Inches(h))
    if caption:
        add_text(slide, x-0.05, y+h+0.03, w+0.1, 0.24,
                 [(caption, dict(size=cap_sz, color=C["mute"], italic=True))],
                 align=PP_ALIGN.CENTER)

def img_ar(img_path):
    """Width/height ratio of an image (needs Pillow)."""
    from PIL import Image
    im = Image.open(img_path); return im.size[0] / im.size[1]

def place_fig(slide, img_path, x, y, w, caption=None, cap_sz=10.5):
    """Framed figure sized to a TARGET WIDTH — height derived from the image's aspect
    ratio so paper crops never get squashed. Returns the rendered height (inches)."""
    h = w / img_ar(img_path)
    figure(slide, img_path, x, y, w, h, caption, cap_sz)
    return h

def hi(slide, x, y, w, h, col=C["good"]):
    """Translucent highlight band (accent border + ~8% alpha fill) over a table row."""
    sp = rrect(slide, x, y, w, h, fill=None, line=col, lw=1.25, radius=0.02)
    sp.fill.solid(); sp.fill.fore_color.rgb = rgb(col)
    srgb = sp.fill.fore_color._xFill.find(qn('a:srgbClr'))
    srgb.append(srgb.makeelement(qn('a:alpha'), {'val': '14000'}))   # ~8% opacity
    return sp

def hi_row(slide, img_path, fx, fy, fw, center_frac, col=C["good"], bandh=0.17, xinset=0.06):
    """Highlight one row inside a placed table image. `center_frac` = the row's vertical
    center as a fraction (0..1) of the rendered image height. Measure it from the LEFT
    label column of the crop, and beware sub-header rows (they're easy to hit by mistake)."""
    fh = fw / img_ar(img_path)
    cy = fy + center_frac * fh
    hi(slide, fx + xinset, cy - bandh/2, fw - 2*xinset, bandh, col)

# ---------------------------------------------------------------- template finders
def _txt(sh):
    return sh.text_frame.text.strip() if sh.has_text_frame else ""

def find_tag(slide, prefix="ACM MM"):
    """The running tag textbox (defaults to decks whose tag starts 'ACM MM')."""
    for sh in slide.shapes:
        if _txt(sh).startswith(prefix): return sh
    return None

def find_footer(slide, prefix=None):
    """The author/footer line. Pass the leading author surname as `prefix`."""
    for sh in slide.shapes:
        t = _txt(sh)
        if prefix and t.startswith(prefix): return sh
        if prefix is None and ("http" in t or "Link" in t) and len(t) > 40: return sh
    return None

def find_pagenum(slide):
    for sh in slide.shapes:
        if _txt(sh).isdigit() and abs(Emu(sh.left).inches - FRAME["pagenum"]["x"]) < 0.3:
            return sh
    return None

def logos(slide):
    return [sh for sh in slide.shapes
            if sh.shape_type == 13 and abs(Emu(sh.top).inches - 0.20) < 0.06]

def strip_to_template(slide, pagenum_str, tag_prefix="ACM MM", footer_prefix=None,
                      also_keep=None, readd_marker=True):
    """Delete everything on a slide except the template chrome
    (tag, footer, page number, logos, + any shapes in `also_keep`).
    Sets the page number and (optionally) re-adds the navy tag marker square.
    Use this to turn a duplicated template slide into a blank-but-branded canvas."""
    tag = find_tag(slide, tag_prefix); ft = find_footer(slide, footer_prefix)
    pn = find_pagenum(slide)
    keep = {sh.shape_id for sh in (tag, ft, pn, *logos(slide)) if sh is not None}
    if also_keep:
        keep |= {sh.shape_id for sh in also_keep if sh is not None}
    for sh in list(slide.shapes):
        if sh.shape_id not in keep:
            sh._element.getparent().remove(sh._element)
    if pn is not None:
        for r in pn.text_frame.paragraphs[0].runs: r.text = pagenum_str
    if readd_marker:
        m = FRAME["tag_marker"]
        rrect(slide, m["x"], m["y"], m["w"], m["h"], fill=C["navy"], radius=0.3)
    return slide

def interior_header(slide, kicker, title_cn):
    """Standard interior-page header: small blue kicker + large navy title."""
    add_text(slide, FRAME["kicker"]["x"], FRAME["kicker"]["y"], 9.0, 0.3,
             [(kicker.upper(), dict(size=FRAME["kicker"]["sz"], color=C["blue"], bold=True, name=EN))])
    add_text(slide, FRAME["title"]["x"], FRAME["title"]["y"], 11.5, 0.6,
             [(title_cn, dict(size=FRAME["title"]["sz"], color=C["navy"], bold=True))])


# ================================================================ embedded logos
# The two header logos are embedded here (base64 of palette-optimized PNGs) so this
# single file is fully self-contained — no separate asset files need to persist.
# NOTE: filenames are historical. On the slide the LEFT header slot (FRAME["logo_nwpu"],
# x=9.101) shows the ASLP@NPU lab mark; the RIGHT slot (FRAME["logo_aslp"], x=10.555)
# shows the NWPU university mark. The file/key names are inverted vs. their content but
# are self-consistent, so rendering is correct. Don't "fix" the names without also
# swapping the FRAME coords.
import base64 as _b64, os as _os

_LOGO_B64 = {
  "logo_aslp.png": "",   # (base64 omitted in SKILL.md; real logos ship in the .skill bundle)
  "logo_nwpu.png": "",
}

def write_logos(dirpath="assets"):
    """Materialize the two header logos into `dirpath` (no-op if none are embedded).
    Returns the dirpath. Call before building so figure()/add_picture() can find them."""
    _os.makedirs(dirpath, exist_ok=True)
    for name, b64 in _LOGO_B64.items():
        data = "".join(b64.split())
        if not data:
            continue
        with open(_os.path.join(dirpath, name), "wb") as f:
            f.write(_b64.b64decode(data))
    return dirpath

def has_logos():
    """True if real logo bitmaps are embedded in this module."""
    return any("".join(v.split()) for v in _LOGO_B64.values())

def logo_path(which, dirpath="assets"):
    """Return path to a logo, writing it first if missing. which in {'aslp','nwpu'}."""
    name = "logo_%s.png" % which
    p = _os.path.join(dirpath, name)
    if not _os.path.exists(p):
        write_logos(dirpath)
    return p if _os.path.exists(p) else None

def header_logos(slide, dirpath="assets"):
    """Place the two top-right header marks. Uses the embedded logo bitmaps when
    available; otherwise draws clean TEXT brand-marks so the header never breaks.
    LEFT slot = ASLP@NPU lab mark, RIGHT slot = NWPU university mark."""
    if has_logos():
        from pptx.util import Inches as _In
        for key in ("logo_nwpu", "logo_aslp"):   # file names are historical (see note above)
            f = FRAME[key]; p = logo_path(key.split("_")[1], dirpath)
            if p:
                slide.shapes.add_picture(p, _In(f["x"]), _In(f["y"]), _In(f["w"]), _In(f["h"]))
        return
    # text fallback (no bitmaps embedded / bundle not installed)
    add_text(slide, 8.55, 0.20, 2.0, 0.44,
             [("ASLP", dict(size=17, color=C["navy"], bold=True, name=EN)),
              ("@NPU", dict(size=17, color=C["blue"], bold=True, name=EN))],
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    add_text(slide, 10.75, 0.20, 1.98, 0.44,
             [("NWPU", dict(size=17, color=C["navy"], bold=True, name=EN))],
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    # thin divider between the two marks
    rrect(slide, 10.62, 0.26, Pt(1.0)/Inches(1), 0.32, fill=C["line"])

if __name__ == "__main__":
    # smoke test: palette + font patch produce valid objects
    print("palette keys:", len(C), "| CN:", CN, "| EN:", EN)
    print("FRAME keys:", list(FRAME.keys()))
    import tempfile, os as _o
    d = write_logos(tempfile.mkdtemp())
    print("logos written:", sorted(_o.listdir(d)))
    print("pptx_cards loaded OK (self-contained)")

```
