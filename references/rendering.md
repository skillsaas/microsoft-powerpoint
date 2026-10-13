# Rendering: PDF, PNG, notes pages, legacy conversion

`pptx_render.py` is the only route this skill uses to run LibreOffice on a deck.

## Why the sandbox
*Tested* with LibreOffice 25.2.3.2: a deck with one picture linked to `http://127.0.0.1:PORT/...`
and one linked to `file:///tmp/.../secret.png`, converted with a plain
`soffice --headless --convert-to pdf`, made an HTTP request to the local server and embedded
the local file in the PDF. The same deck through `pptx_render.py` made no request and the
PDF had no images (`pdfimages -list`). Both cases are smoke tests.

## What the sandbox is
bubblewrap with `--unshare-all` (no network, new PID/IPC/UTS/user namespaces), read-only `/usr`,
fonts and the linker cache, `/etc/libreoffice`, empty `/tmp` and `/run`, and one work folder
with a private copy of the input and a fresh LibreOffice profile (macro security very high,
macro execution disabled, untrusted referer links blocked). No home folder, no other user files,
minimal environment, process group killed at `--timeout`. Public files under `/usr` remain visible.

## Commands and tested behavior
```bash
python <skill-dir>/scripts/pptx_render.py --check --json
python <skill-dir>/scripts/pptx_render.py deck.pptx deck.pdf --json
python <skill-dir>/scripts/pptx_render.py deck.pptx deck-notes.pdf --notes-pages
python <skill-dir>/scripts/pptx_render.py deck.pptx deck-png --to png --dpi 110
python <skill-dir>/scripts/pptx_render.py old.ppt old.pptx --to pptx
```
- Hidden slides were left out of the PDF (3 visible of 4). PNG files follow PDF pages, so
  `slide-2.png` is the second *visible* slide, not slide 2 of the deck. Map with
  `inspect_pptx.py` (`hidden`).
- `--notes-pages` uses the PDF filter options `ExportNotesPages` and `ExportOnlyNotesPages`
  (JSON filter syntax); the PDF had one notes page per visible slide with the notes text.
- `soffice --convert-to png` alone wrote only the first slide (`three.png` for a 3-slide deck),
  which is why PNG goes through PDF and pdftoppm.
- pdftoppm numbers files `slide-1.png ...`, zero-padded when there are 10 or more pages
  (poppler behavior, documented, not tested here).
- `.ppt` to `.pptx`: LibreOffice re-saves the whole deck (layout names became "Title7" in our
  test, metadata changed). Keep the original; tell the user it was converted.
- Encrypted decks are refused before LibreOffice runs.

## Fidelity: LibreOffice is not PowerPoint
Expect differences in: fonts not installed in the sandbox (substitution changes line breaks
and autofit), SmartArt, some chart types,
3D models, morph and other effects, gradients and shadows, embedded fonts (all documented as
known LibreOffice gaps, not each tested here). Always say "rendered
with LibreOffice" when you show images, and when exact appearance matters ask the user to
export from PowerPoint or use Graph `?format=pdf` (see `automation.md`).
To check text overflow, compare the rendered images with the text from `inspect_pptx.py`.

## Timeouts and failures
Default 180 s. Exit 3 on timeout (process group killed, nothing written); retry once with
`--timeout 600` for large media decks. Exit 1 when LibreOffice ran and wrote nothing (often a
corrupt or unsupported deck): run `inspect_pptx.py` to see why.
