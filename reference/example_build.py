#!/usr/bin/env python3
"""
example_build.py — worked example for the paper-card-deck skill.

Demonstrates:
  1. building a branded "blank" template slide (chrome only) from assets,
  2. the two-column Cover archetype,
  3. the Audio-Demo archetype (2x2 placeholder cards).

Run from the skill root:  python3 reference/example_build.py
Outputs: example_out.pptx  (+ template_blank.pptx)

The helper library lives in ../scripts; we add it to sys.path so the example is
runnable from the skill root regardless of cwd.
"""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))
ASSETS = os.path.join(ROOT, "assets")

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx_cards import (C, CN, EN, FRAME, rgb, set_font, add_text, rrect, dot,
                        figure, interior_header, header_logos, find_pagenum)

# ----------------------------------------------------------------- chrome builder
def new_pres():
    p = Presentation(); p.slide_width = Inches(13.333); p.slide_height = Inches(7.5)
    return p

def branded_slide(prs, pagenum, tag_text, footer_text):
    """A blank slide carrying the shared chrome: tag+marker, logos, pagenum, footer."""
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    # running tag + navy marker
    m = FRAME["tag_marker"]; rrect(s, m["x"], m["y"], m["w"], m["h"], fill=C["navy"], radius=0.3)
    add_text(s, FRAME["tag"]["x"], FRAME["tag"]["y"], FRAME["tag"]["w"], FRAME["tag"]["h"],
             tag_text, anchor=MSO_ANCHOR.MIDDLE)
    # logos (real bitmaps if embedded/available, else text brand-marks)
    header_logos(s, os.path.join(ROOT, "assets"))
    # page number
    f = FRAME["pagenum"]
    add_text(s, f["x"], f["y"], f["w"], f["h"],
             [(str(pagenum), dict(size=11, color=C["faint"], name=EN))], align=PP_ALIGN.RIGHT)
    # footer
    add_text(s, 0.49, 6.96, 11.9, 0.47, footer_text, line_spacing=1.05)
    return s

TAG = [("ACM MM 2026  ·  ", dict(size=9.5, color=C["tagmute"], name=EN)),
       ("MINT-Bench", dict(size=9.5, color=C["navy"], bold=True, name=EN)),
       ("：面向指令跟随 TTS 的多语种评测基准", dict(size=9.5, color=C["tagmute"]))]
FOOTER = [("Author A, Author B, … , Lei Xie. ACM MM 2026 Oral. ", dict(size=10, color=C["mute"])),
          ("Paper Link: ", dict(size=10, color=C["ink"], bold=True)),
          ("arxiv.org/abs/XXXX", dict(size=10, color=C["blue"], name=EN, link="https://arxiv.org")),
          (" ; Demo Link: ", dict(size=10, color=C["ink"], bold=True)),
          ("demo.site", dict(size=10, color=C["blue"], name=EN, link="https://example.com")),
          (" ; GitHub Link: ", dict(size=10, color=C["ink"], bold=True)),
          ("github.com/org/repo", dict(size=10, color=C["blue"], name=EN, link="https://github.com"))]

# ----------------------------------------------------------------- cover archetype
def cover(prs):
    s = branded_slide(prs, 1, TAG, FOOTER)
    add_text(s, 0.55, 0.52, 12.2, 0.7,
             [("MINT-Bench", dict(size=27, color=C["navy"], bold=True, name=EN)),
              ("：面向指令跟随 TTS 的综合性多语种评测基准", dict(size=27, color=C["ink"], bold=True))])
    add_text(s, 0.55, 1.22, 12.1, 0.34,
             [("A Comprehensive Multilingual Benchmark for Instruction-Following Text-to-Speech",
               dict(size=13.5, color=C["mute"], italic=True, name=EN))])
    dv = rrect(s, 0.55, 1.66, 12.23, Pt(1.1)/Inches(1), fill=C["line"])  # hairline
    # LEFT — background bullets
    add_text(s, 0.55, 1.86, 5.75, 0.3, [("研究背景与动机 · Background", dict(size=14, color=C["navy"], bold=True))])
    bl = [("控制范式升级", C["tax"], "指令跟随 TTS 正从「读对内容」走向「用自然语言指令控制 音色/风格/情感/角色」。"),
          ("评测过于扁平", C["data"], "现有基准多为扁平 prompt 集合,缺乏系统化难度分层与细粒度可诊断性。"),
          ("维度易混淆", C["eval"], "多语种覆盖不足,且常将「内容一致性」与「指令跟随」混为一谈。"),
          ("研究动机", C["warn"], "亟需一个统一、可诊断、可扩展的指令跟随合成评测框架。")]
    y = 2.34
    for lead, col, rest in bl:
        rrect(s, 0.55, y+0.055, 0.12, 0.12, fill=col, radius=0.3)
        add_text(s, 0.85, y, 5.45, 0.9,
                 [("【"+lead+"】", dict(size=12, color=col, bold=True)),
                  (rest, dict(size=12, color=C["mute"]))], line_spacing=1.14)
        y += 1.00
    # RIGHT — approach: figure + notes (figure omitted here if asset missing)
    add_text(s, 6.55, 1.86, 6.18, 0.3, [("提出方案 · MINT-Bench", dict(size=14, color=C["navy"], bold=True))])
    fig = os.path.join(ASSETS, "fig1_overview.png")
    if os.path.exists(fig):
        figure(s, fig, 6.55, 2.30, 6.18, 6.18*0.42, "图 1 · 总体框架")
    else:
        rrect(s, 6.55, 2.30, 6.18, 2.60, fill=C["panel"], line=C["line"], radius=0.03)
        add_text(s, 6.55, 3.4, 6.18, 0.4, [("[ 论文框架图占位 · 用 figure() 插入裁剪图 ]",
                 dict(size=12, color=C["faint"]))], align=PP_ALIGN.CENTER)
    notes = [(C["tax"], "Taxonomy-driven Benchmark", "层级多轴组织评测空间"),
             (C["data"], "Scalable Data Construction", "taxonomy node 三阶段构造"),
             (C["eval"], "Hierarchical Hybrid Evaluation", "content → instruction → perceptual")]
    ny = 5.12
    for col, en, cn in notes:
        rrect(s, 6.55, ny+0.055, 0.12, 0.12, fill=col, radius=0.3)
        add_text(s, 6.85, ny, 5.9, 0.32,
                 [(en+"  ", dict(size=11.5, color=C["navy"], bold=True, name=EN)),
                  ("· "+cn, dict(size=11, color=C["mute"]))])
        ny += 0.40
    return s

# ----------------------------------------------------------------- audio-demo archetype
def audio_demo(prs):
    s = branded_slide(prs, 5, TAG, FOOTER)
    interior_header(s, "04 · AUDIO DEMO", "指令跟随合成 · 试听样例")
    add_text(s, 0.55, 1.30, 12.15, 0.3,
             [("每个样例给出指令与合成文本,音频 / 二维码位置已预留 —— 后续替换为实际 demo 即可。",
               dict(size=12.5, color=C["mute"]))])
    demo = [("Timbre · Age = Child", "Easy", C["tax"], "Speak as if you are a young child.", "Can I have a glass of juice, please?"),
            ("Style · Emotion → Dynamic", "Hard", C["data"], "Start calm, then gradually become excited.", "I think… wait — we actually won!"),
            ("Composition · Layered Control", "Hard", C["eval"], "An elderly man, slow, warm, gentle.", "Take your time, child."),
            ("Special · Disfluency", "Special", C["warn"], "Speak with occasional hesitations and fillers.", "So, um… let's just try it.")]
    CW, CH = 6.0, 2.52; xs = [0.55, 6.78]; ys = [1.72, 4.40]
    for i, (cat, lvl, col, instr, text) in enumerate(demo):
        cx, cy = xs[i % 2], ys[i // 2]
        rrect(s, cx, cy, CW, CH, fill=C["panel"], line=C["line"], radius=0.06)
        chip = rrect(s, cx+0.24, cy+0.22, 3.0, 0.44, fill=col, radius=0.3)
        tf = chip.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = cat; set_font(r, size=12.5, color=C["white"], bold=True, name=EN)
        add_text(s, cx+CW-1.7, cy+0.27, 1.46, 0.34, [(lvl, dict(size=12.5, color=col, bold=True, name=EN))], align=PP_ALIGN.RIGHT)
        add_text(s, cx+0.24, cy+0.80, CW-0.48, 0.24, [("INSTRUCTION", dict(size=10.5, color=C["faint"], bold=True, name=EN))])
        add_text(s, cx+0.24, cy+1.06, CW-0.48, 0.3, [("“"+instr+"”", dict(size=13, color=C["ink"], name=EN))])
        add_text(s, cx+0.24, cy+1.46, CW-0.48, 0.24, [("SYNTHESIS TEXT", dict(size=10.5, color=C["faint"], bold=True, name=EN))])
        add_text(s, cx+0.24, cy+1.72, CW-0.48, 0.3, [("“"+text+"”", dict(size=13, color=C["mute"], italic=True, name=EN))])
        ph = rrect(s, cx+0.24, cy+2.08, CW-0.48, 0.40, fill=C["white"], line=col, lw=1.25, radius=0.08, dash="dash")
        tf = ph.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = "▶  音频 / 二维码占位 · 在此插入试听"; set_font(r, size=12, color=col, bold=True)
    return s

if __name__ == "__main__":
    prs = new_pres()
    cover(prs)
    audio_demo(prs)
    out = os.path.join(ROOT, "example_out.pptx")
    prs.save(out)
    print("wrote", out, "| slides:", len(prs.slides._sldIdLst))
