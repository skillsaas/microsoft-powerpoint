# microsoft-powerpoint

An agent skill (`SKILL.md`) that teaches AI coding agents to work with Microsoft PowerPoint
decks safely: read them, edit text without losing formatting, render previews, and clean them
before they are shared.

## Use it when
- a user asks you to read, summarize, edit or rebrand a `.pptx`, `.pptm`, `.ppsx` or `.potx`;
- find and replace misses words because PowerPoint split them across runs;
- python-pptx cannot delete, reorder or hide a slide, or `text_frame.text` wiped formatting;
- you need slide images or a PDF (LibreOffice in a sandbox), or notes pages as a handout;
- a chart's numbers must change and its embedded workbook must follow;
- a deck came from outside (macros, OLE objects, linked pictures, click actions that run programs);
- a deck is about to be sent and may still hold speaker notes, comments, hidden slides or metadata;
- you automate PowerPoint with VBA, an Office.js add-in or Microsoft Graph.

## Install
Planned, not yet published (v0.1.0 is under review):
```bash
npx @skillsaas/cli add microsoft-powerpoint      # planned
npx skills add skillsaas/microsoft-powerpoint    # planned
```
Today: copy this folder into your agent's skills folder (for example `.claude/skills/`
or `.agents/skills/`) and run `pip install -r scripts/requirements.txt`.

## Contents
| Path | What |
| :--- | :--- |
| `SKILL.md` | Router: decision guide, workflow, rules, exit codes, checklist |
| `references/` | PresentationML, python-pptx, charts and media, rendering, security, recipes, extraction, automation, fallbacks, troubleshooting |
| `scripts/inspect_pptx.py` | Read-only inventory and risk list |
| `scripts/pptx_replace.py` | Replace text split across runs, with `--expect` |
| `scripts/pptx_sanitize.py` | Share-check and share-ready copy |
| `scripts/pptx_render.py` | PDF / PNG / legacy conversion with LibreOffice in bubblewrap |
| `tests/smoke_test.py` | Tests, run with `python tests/smoke_test.py` |

Requirements: Python 3.9+ and lxml. Rendering needs Linux with LibreOffice, bubblewrap and
poppler-utils. See the catalog at https://github.com/skillsaas/index.

## Trademark
Unofficial. Not affiliated with, endorsed by or sponsored by Microsoft. Microsoft and
PowerPoint are trademarks of the Microsoft group of companies.

License: MIT.
