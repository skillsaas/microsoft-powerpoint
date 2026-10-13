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
| Deleted slide text still found in the file | `unlisted_slide_parts` in inspect | `pptx_sanitize.py --remove-hidden-slides`; delete with `pptx_slides.py --delete` next time |
| Chart shows new numbers but "Edit Data" shows old | Was only the XML cache edited? | `chart.replace_data()` |
| Video does not play on another computer | `media[].linked == true` | Embed the media (Insert > Video > This Device) or ship the file with the deck |
| PNG count differs from slide count | Hidden slides | Expected: renders skip hidden slides; map with `inspect_pptx.py` |
| Rendered fonts look wrong | Fonts missing in the sandbox | Install the fonts system-wide (they are bound from `/usr` and `/etc/fonts`), or export from PowerPoint |
| `pptx_render.py` exit 2 "cannot sandbox" | `--check --json` `bwrap_error` | See `fallbacks.md`; do not render unsandboxed |
| Exit 2 "password-protected" | | Ask for an unprotected copy |
| Exit 3 "duplicate zip entry" or caps | | Do not process; tell the user the file looks malformed |
| `--check-clean` still fails after `--all` | `remaining` | Report each item; something the scripts cannot remove needs PowerPoint (Inspect Presentation) |
| A .pptm saved as .pptx will not open | Main content type still macro-enabled | `pptx_sanitize.py --remove-macros` writes the right type and requires the .pptx extension |
| A link stopped working after deleting a slide | `pptx_slides.py` `changes`: "link to removed slide" | Expected: links to a deleted slide are removed, text kept. Relink in PowerPoint if wanted |
| `pptx_build.py` exit 1 "no placeholder idx" | `--list-layouts TEMPLATE` | Use an idx that layout has, or another layout |
| Built slides lack date, footer, slide number | Layout has idx 10 to 12, slide does not | python-pptx does not copy them; Insert > Header & Footer in PowerPoint |
| `mismatched-link` on a link the user trusts | `links[]` `mismatch` and `text` (`--show-targets` for the target) | Ask; keep it by not using `--remove-active`, or fix the text to show the real site |
| Render exit 1 or 3 on a big deck | `--memory-mb`, `--timeout` | Retry once with `--memory-mb 8192 --timeout 600`; otherwise export from PowerPoint |
| "OUT must end in .pptx/.pptm" | Package type vs extension | Use the extension the error names; `--remove-macros` to go from .pptm to .pptx |
