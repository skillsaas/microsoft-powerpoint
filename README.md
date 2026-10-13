# microsoft-powerpoint

**TL;DR:** `microsoft-powerpoint` is a free, MIT-licensed agent skill that teaches AI agents (Claude Code, Cursor, Codex, Gemini CLI and other agents that read SKILL.md skills) to read, edit and build PowerPoint decks without losing formatting, render slide previews in a sandbox, and clean decks before they are shared. Install: `npx skills add skillsaas/microsoft-powerpoint` or `npm install microsoft-powerpoint`.

An agent skill (`SKILL.md`) that teaches AI coding agents to work with Microsoft PowerPoint
decks safely: read them, edit text without losing formatting, render previews, and clean them
before they are shared.

## Use it when
- a user asks you to read, summarize, edit or rebrand a `.pptx`, `.pptm`, `.ppsx` or `.potx`;
- find and replace misses words because PowerPoint split them across runs;
- slides must be deleted, reordered or hidden (python-pptx has no API), or `text_frame.text` wiped formatting;
- a new deck must be built from the user's template, filling layout placeholders;
- you need slide images or a PDF (LibreOffice in a sandbox), or notes pages as a handout;
- a chart's numbers must change and its embedded workbook must follow;
- a deck came from outside (macros, OLE objects, linked pictures, click actions that run programs);
- a deck is about to be sent and may still hold speaker notes, comments, hidden slides or metadata;
- you automate PowerPoint with VBA, an Office.js add-in or Microsoft Graph.

## Install
```bash
npx skills add skillsaas/microsoft-powerpoint    # skills CLI, reads SKILL.md at the repo root
npm install microsoft-powerpoint                 # files land in node_modules/microsoft-powerpoint/
# Planned, not yet published: npx @skillsaas/cli add microsoft-powerpoint
```
Or copy this folder into your agent's skills folder (for example `.claude/skills/`
or `.agents/skills/`), then run `pip install -r scripts/requirements.txt`.

Try it: `python scripts/inspect_pptx.py tests/fixtures/deck.pptx`, then `python tests/smoke_test.py`.

## What it does
- `inspect_pptx.py`: read-only inventory and risk list (macros, OLE objects, linked pictures, click actions that run programs, notes, comments, hidden slides).
- `pptx_replace.py`: replaces text split across runs, with an `--expect` count.
- `pptx_slides.py`: deletes, hides, unhides and moves slides completely (python-pptx has no API for this).
- `pptx_build.py`: builds a new deck from a `.pptx` or `.potx` template by placeholder (needs python-pptx), with `--dry-run`.
- `pptx_sanitize.py` share-check and share-ready copy; `pptx_render.py` PDF or PNG with LibreOffice in bubblewrap (Linux).
- Python 3.9+ with lxml; references for python-pptx, PresentationML, charts, VBA, Office.js and Microsoft Graph.

## FAQ
**What is the microsoft-powerpoint skill?**
An Agent Skills package (a `SKILL.md` plus references and Python scripts) that tells an AI agent how to work with PowerPoint decks correctly. It is unofficial and not affiliated with Microsoft.

**How do I delete or reorder slides with python-pptx?**
python-pptx has no API for it. `pptx_slides.py` deletes, hides and moves slides and removes the parts a deleted slide leaves behind.

**Why does replacing text in a .pptx lose formatting?**
Setting `text_frame.text` rewrites the runs. `pptx_replace.py` replaces text across runs and keeps their formatting.

**How do I make slide images or a PDF without PowerPoint?**
`pptx_render.py` uses LibreOffice in a bubblewrap sandbox on Linux; PowerPoint remains the reference for exact visual fidelity.

**How do I install it in Claude Code?**
`npx skills add skillsaas/microsoft-powerpoint`, or copy the folder to `~/.claude/skills/microsoft-powerpoint/`, then `pip install -r scripts/requirements.txt`.

**How do I clean a deck before sending it?**
`pptx_sanitize.py` lists speaker notes, comments, hidden slides and metadata and writes a share-ready copy.

## Contents
| Path | What |
| :--- | :--- |
| `SKILL.md` | Router: decision guide, workflow, rules, exit codes, checklist |
| `references/` | PresentationML, python-pptx, charts and media, rendering, security, recipes, extraction, automation, fallbacks, troubleshooting |
| `scripts/inspect_pptx.py` | Read-only inventory and risk list |
| `scripts/pptx_replace.py` | Replace text split across runs, with `--expect` |
| `scripts/pptx_slides.py` | Delete, hide, unhide, move slides completely |
| `scripts/pptx_build.py` | New deck from a .pptx or .potx template by placeholder idx (needs python-pptx) |
| `scripts/pptlinks.py`, `scripts/pptpkg.py` | Shared helpers: link checks, safe package reader |
| `scripts/pptx_sanitize.py` | Share-check and share-ready copy |
| `scripts/pptx_render.py` | PDF / PNG / legacy conversion with LibreOffice in bubblewrap |
| `tests/smoke_test.py` | Tests, run with `python tests/smoke_test.py` |

Requirements: Python 3.9+ and lxml>=5 (tested 6.1.3); python-pptx 1.0.2 for `pptx_build.py`. Rendering needs Linux with LibreOffice, bubblewrap and
poppler-utils. See the catalog at https://github.com/skillsaas/index.

## Status and maintenance

Version 0.1.4, published on npm and GitHub (pre-1.0). Smoke tests run on Linux only; not run on macOS or Windows. Each version is
reviewed by a panel of people and AI agents before the next; CHANGELOG.md lists what changed and
why. Versions follow semver; until 1.0, a minor bump may change script flags or JSON fields, and
the CHANGELOG says so. Run `python tests/smoke_test.py` after install to check your setup.

## Trademark
Unofficial. Not affiliated with, endorsed by or sponsored by Microsoft. Microsoft and
PowerPoint are trademarks of the Microsoft group of companies.

License: MIT.
