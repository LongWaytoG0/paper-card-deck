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
  "logo_aslp.png": (
"""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"""
  ),
  "logo_nwpu.png": (
"""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
QmCC"""
  ),
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
