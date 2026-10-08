# AGENTS.md — Paper-Card Deck skill

This repository is a **self-contained skill**: it turns a research paper (PDF + links) into an
ASLP@NPU lab-intro "paper-card" slide deck — a **simple version (2–3 pages)** and a
**detailed version (3–6 pages)** on one shared branded PowerPoint template.

**The full method, page archetypes, coordinates, and the helper-library source are in
[`SKILL.md`](./SKILL.md). Read it before building.** This file is the short agent-facing entry
point; `SKILL.md` is the spec.

## When to use
The user gives you a paper (a PDF, usually plus an arXiv/paper link and maybe demo/GitHub
links) and wants lab-style introduction slides / "paper cards" / a module to merge into a
combined lab deck. Output = two `.pptx` files (`<Name>_详细版.pptx` and `<Name>_简版.pptx`).

## Files in this skill
- `SKILL.md` — the complete instructions (workflow, page archetypes with coordinates, QA).
- `scripts/pptx_cards.py` — the helper library. **Import it directly** (`from pptx_cards import *`);
  you do NOT need to recreate it from SKILL.md's Appendix A — that embedded copy exists only so
  the single file can stand alone when a host can't ship sibling files. In this repo the real
  file (with the embedded brand logos) is already on disk, so just import it.
- `assets/logo_aslp.png`, `assets/logo_nwpu.png` — the two header brand marks.
- `reference/page_archetypes.md`, `reference/example_build.py` — a worked example deck.
- `requirements.txt` — `python-pptx` + `Pillow` (plus optional LibreOffice/poppler for QA).

## Setup
```bash
pip install -r requirements.txt
```
Then build with Python using `scripts/pptx_cards.py` and the archetypes in `SKILL.md`.
`reference/example_build.py` is a complete, runnable example to copy from.

## Core conventions — DO NOT VIOLATE (full detail in SKILL.md)
1. **Reuse the paper's own figures & tables as cropped images** + Chinese explanation; do not
   re-typeset dense result tables natively.
2. **中文讲解 + 英文术语混排** — narrate in Chinese, keep English technical terms.
3. **2–6 pages per module**, never a long talk. Detailed default = 5p (4 paper + 1 demo);
   simple = a 4p subset that **reuses the same page functions** (don't rebuild).
4. **Preserve the template chrome exactly** on every page (logos top-right, running tag
   top-left, page number bottom-right, author+links footer bottom-left).
5. **Never regenerate-and-overwrite a file the user may have hand-edited** — edit in place and
   back it up first.
6. **Only show links the user actually provided — never fabricate or repurpose one.** No demo
   link ⇒ no demo link and no "online-demo / 在线试听" pill. Don't dress up a leaderboard or
   project page as a demo. Omit absent links; a clearly-labeled placeholder box is fine.

## Validate
Render every page (`soffice --headless --convert-to pdf …` then `pdftoppm -jpeg …`) and eyeball
the JPEGs before delivering. Note: Linux renderers often lack 微软雅黑 and substitute a wider
face, so CJK titles may *appear* to wrap in the preview but are fine in real PowerPoint.
