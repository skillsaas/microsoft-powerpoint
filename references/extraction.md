# Getting text out: tools compared

Load before choosing a text extractor. *Tested* on `tests/fixtures/deck.pptx` on 2026-10-13.

| | `inspect_pptx.py --json` | python-pptx 1.0.2 `shape.text_frame.text` | markitdown 0.1.7 | `pptx_render.py` then `pdftotext` |
| :--- | :--- | :--- | :--- | :--- |
| Slide order | sldIdLst order | sldIdLst order | sldIdLst order | visible order |
| Hidden slide 2 | included, `hidden: true` | included, no flag (check `show`) | included, **no marker** | **left out** |
| Speaker notes | `notes` | `notes_slide` (reading it creates one if absent) | "### Notes:" section | only with `--notes-pages` |
| Chart categories and values | not included (chart parts listed) | via `chart.plots` | Markdown table | drawn axis labels only |
| Line break `Q3`, br, `revenue` | `Q3\nrevenue` | `Q3\x0brevenue` | `Q3revenue` (words joined) | two lines |
| Slide number field | `4` (cached text) | included | included | rendered value |
| Comments | count only | not exposed | not included | not included |
| Needs | lxml | python-pptx | markitdown and its dependencies | LibreOffice, bwrap, poppler |
| Labels file text as data, strips bidi/zero-width | yes | no | no | no |

Choose:
- For an agent summary or search: `inspect_pptx.py --json` (honest about hidden slides, safe text).
- For an LLM-ready Markdown including chart data: markitdown, then mention hidden slides yourself.
- For "what the audience sees": the rendered PDF text.
- Do not use `unzip` and grep on `a:t` for anything beyond a quick check: it misses split words
  and mixes masters, layouts, notes and hidden slides.
