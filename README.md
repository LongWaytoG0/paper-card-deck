# Paper-Card Deck · ASLP@NPU lab-intro slides

Turn a research paper (PDF + links) into a branded **"paper-card" slide deck** — a compact
**simple version (2–3 pages)** and a **detailed version (3–6 pages)** on one shared template,
so modules from different papers look identical and can be merged into one combined lab deck.

Built for the ASLP@NPU (NWPU) lab style: Chinese narration + English technical terms, the
paper's own figures/tables reused as crops with highlight callouts, consistent header/footer
chrome. The default detailed layout is 5 pages — ①封面（背景+方案图）②方法总览 ③方法细节（按论文灵活
定制）④结果 ⑤demo — and the simple version reuses the same pages as a 4-page subset.

---

## What's in here

```
paper-card-deck/
├── README.md                 ← you are here (human setup + sharing + cross-model notes)
├── SKILL.md                  ← the full skill spec (method, page archetypes, QA, helper source)
├── AGENTS.md                 ← short agent entry point (Codex / Cursor / generic harnesses)
├── requirements.txt          ← python-pptx + Pillow
├── scripts/
│   └── pptx_cards.py         ← the helper library (import this; real brand logos embedded)
├── assets/
│   ├── logo_aslp.png         ← ASLP@NPU lab mark
│   └── logo_nwpu.png         ← NWPU university mark
└── reference/
    ├── page_archetypes.md    ← quick coordinate cheat-sheet
    └── example_build.py      ← a complete, runnable example deck to copy from
```

Everything is plain Markdown + Python — **no proprietary format**. `SKILL.md` and `AGENTS.md`
hold the same instructions in the two conventions different tools expect.

### Dependencies
```bash
pip install -r requirements.txt     # python-pptx, Pillow
# optional, for the render-and-eyeball QA step:
#   LibreOffice (soffice) + poppler-utils (pdftoppm)  — install via your OS package manager
```

---

## First, the key idea: skills belong to the *harness*, not the *model*

A "skill" is just **instructions + helper files** that an agent tool loads into context. Which
file a tool looks for differs:

| Tool / harness                     | Reads this entry file        |
|-----------------------------------|------------------------------|
| Claude Code, Claude apps          | `SKILL.md` (YAML frontmatter) |
| OpenAI **Codex** CLI              | `AGENTS.md`                   |
| Cursor / Cline / Roo Code / Aider | `AGENTS.md` (or project rules)|
| Your own agent script (any API)   | load `SKILL.md` as the system prompt |

**DeepSeek and Kimi are models, not harnesses** — by themselves they have no "skill folder".
You use this skill with them by running them *inside* one of the harnesses above (most let you
point the model endpoint at DeepSeek/Kimi), **or** by pasting `SKILL.md` into the system prompt
and giving the model a code sandbox that has `scripts/pptx_cards.py` + `assets/`. See the
per-tool steps below.

---

## How to install / use

### Claude Code (or Claude desktop skills)
`SKILL.md` already has the right frontmatter. Put the folder where Claude Code finds skills:
```bash
# personal skill
cp -r paper-card-deck ~/.claude/skills/
# or a project skill
cp -r paper-card-deck <your-repo>/.claude/skills/
```
Then invoke it by name (`paper-card-deck`) or just describe the task with a paper PDF attached.
The sibling `scripts/` and `assets/` travel with the skill directory.

### Codex CLI
Codex reads `AGENTS.md`. Either drop this folder into your project root (Codex will pick up
`AGENTS.md`), or copy its contents into your project's top-level `AGENTS.md`:
```bash
cp -r paper-card-deck <your-project>/            # keeps AGENTS.md at a path Codex scans
```
Make sure `pip install -r requirements.txt` has run in the environment Codex executes in.

### Cursor / Cline / Roo / Aider (works with ANY model incl. DeepSeek & Kimi)
These harnesses read `AGENTS.md` (or a rules file) and let you choose the model/endpoint:
1. Copy this folder into your project so `AGENTS.md`, `scripts/`, `assets/` are present.
2. In the tool's settings, select your model — e.g. a DeepSeek or Kimi (Moonshot) API endpoint.
3. Ask it to build paper cards from your PDF; it will follow `AGENTS.md` → `SKILL.md`.

### Any model via raw API or chat UI (DeepSeek, Kimi, GPT, Qwen, …)
No agent framework needed — just supply the instructions and a place to run Python:
1. **System prompt:** paste the entire contents of `SKILL.md` (it is self-contained — Appendix A
   even carries the full helper source so the model can recreate `pptx_cards.py` if it has no
   filesystem).
2. **Files:** if the chat tool has a code interpreter / sandbox, upload `scripts/pptx_cards.py`,
   `assets/logo_aslp.png`, `assets/logo_nwpu.png`, and the paper PDF. If it has no sandbox, the
   model can still emit a complete Python script for you to run locally.
3. Give it the paper PDF + the links you actually have, and ask for the 详细版 + 简版 decks.

> The helper library is pure `python-pptx` and has **no Anthropic/Claude dependency**, so it
> runs the same no matter which model wrote the calling code.

---

## Sharing it with others

**As a Git repo (recommended):**
```bash
cd paper-card-deck
git init && git add . && git commit -m "Paper-card deck skill"
git remote add origin git@github.com:<you>/paper-card-deck.git
git push -u origin main
```
Others clone it and follow this README. Because both `SKILL.md` and `AGENTS.md` are included,
it works out-of-the-box for Claude Code, Codex, and Cursor-style tools.

**As a zip:** just send `paper-card-deck.zip` (provided alongside this folder). The recipient
unzips and follows the install section for their tool.

### Notes for redistribution
- **Logos:** `assets/logo_aslp.png` / `logo_nwpu.png` are the NWPU / ASLP@NPU brand marks.
  They're embedded for this lab's decks — swap them (and the palette/author footer in
  `scripts/pptx_cards.py`) if you adapt this skill for a different group.
- **Fonts:** decks use 微软雅黑 (Microsoft YaHei) for Chinese. It renders correctly in real
  PowerPoint; Linux preview renderers may substitute a wider face (cosmetic only).
- No license file is included — add one (e.g. MIT) before publishing publicly if you want to
  set reuse terms.
