# Troubleshooting

| Symptom | Check | Fix |
| :--- | :--- | :--- |
| Find and replace finds nothing, but the word is on the slide | `pptx_replace.py --dry-run --json`; `runs` > 1 in matches means it was split | Use `pptx_replace.py`, not string replace on XML or runs |
| Replace count lower than expected | `skipped_across_breaks_or_fields`, `not_replaced`, `--notes`, `--masters`, `--slides` | Text on a master or layout needs `--masters` (ask); chart text needs chart data; across a line break edit by hand |
| Count higher than expected | `hidden_slides_touched`, matches list | Ask about hidden slides; restrict with `--slides` |
| Formatting lost after edit | Was `text_frame.text` set? | Rebuild from the original with `pptx_replace.py` |
| Text sits in the top-left corner or in the wrong font | `placeholders[].inherits_from == "none"` | Change the slide to a layout that has that placeholder, or give the shape its own `a:xfrm` |
| Text overflows the box after an edit | Render to PNG and look | Shorten text, or let PowerPoint autofit on next open (LibreOffice preview may differ) |
| "PowerPoint found a problem with content" | `inspect_pptx.py`, then the smoke test `consistent()` on the file | Dangling relationship, missing content type, stale section id, reused slide id; see `presentationml.md` repair list |
| Deleted slide text still found in the file | `unlisted_slide_parts` in inspect | `pptx_sanitize.py --remove-hidden-slides` |
| Chart shows new numbers but "Edit Data" shows old | Was only the XML cache edited? | `chart.replace_data()` |
| Video does not play on another computer | `media[].linked == true` | Embed the media (Insert > Video > This Device) or ship the file with the deck |
| PNG count differs from slide count | Hidden slides | Expected: renders skip hidden slides; map with `inspect_pptx.py` |
| Rendered fonts look wrong | Fonts missing in the sandbox | Install the fonts system-wide (they are bound from `/usr` and `/etc/fonts`), or export from PowerPoint |
| `pptx_render.py` exit 2 "cannot sandbox" | `--check --json` `bwrap_error` | See `fallbacks.md`; do not render unsandboxed |
| Exit 2 "password-protected" | | Ask for an unprotected copy |
| Exit 3 "duplicate zip entry" or caps | | Do not process; tell the user the file looks malformed |
| `--check-clean` still fails after `--all` | `remaining` | Report each item; something the scripts cannot remove needs PowerPoint (Inspect Presentation) |
| A .pptm saved as .pptx will not open | Main content type still macro-enabled | `pptx_sanitize.py --remove-macros` writes the right type and requires the .pptx extension |
