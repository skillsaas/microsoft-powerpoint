# python-pptx: what it does and does not do

Tested with python-pptx 1.0.2 on Python 3.13 unless marked otherwise. Load when building decks
or when python-pptx output surprises you.

## Good at
- Building decks from a template: `Presentation("template.pptx")`, `prs.slide_layouts`,
  `prs.slides.add_slide(layout)`, filling `slide.placeholders[idx]`.
- Text with formatting at run level: `para.add_run()`, `run.font.bold`, `run.font.size`.
- Tables, pictures, auto shapes, charts (`shapes.add_chart`), chart data replacement
  (`chart.replace_data`), speaker notes (`slide.notes_slide.notes_text_frame`).
- Keeping parts it does not understand: a `.pptm` opened and saved kept `ppt/vbaProject.bin` (*tested*).

## Does not do (tested in 1.0.2)
| Missing | Workaround |
| :--- | :--- |
| Delete a slide (`prs.slides` has only `add_slide`, `get`, `index`) | Hide it, then `pptx_sanitize.py IN OUT --remove-hidden-slides` (cleans sections, links, notes). The common `drop_rel` + remove `sldId` recipe leaves the part in the file when another slide links to it, and leaves stale section ids |
| Reorder slides | Move `p:sldId` elements in `prs.slides._sldIdLst` (below) |
| Hide or unhide a slide (no `hidden` property) | `slide._element.set("show", "0")`; delete the attribute to unhide |
| Duplicate a slide | Deep-copy shapes and copy relationships except notes, layout and comments (below); test-render the result |
| Keep formatting when setting text | `text_frame.text = "..."` replaced the runs; a bold run came back with `bold None`. Use `pptx_replace.py`, or edit `run.text` per run |
| Render, lay out, measure text, compute autofit | Render with `pptx_render.py`; check overflow on the images |
| Sections, transitions, animations, comments, media playback settings | XML only (see `presentationml.md`, `charts-media.md`) |
| Open `.ppt`, `.odp` or encrypted files | `pptx_render.py IN OUT.pptx --to pptx`; ask for an unprotected copy |
Also observed: reading `slide.notes_slide` **creates** a notes slide when there is none
(`has_notes_slide` went from False to True). Check `has_notes_slide` first.
The default template (`Presentation()` with no file) is 4:3, 9144000 x 6858000 EMU. Start from
the user's template, or set `prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)` for 16:9.

## Recipes (tested together on the fixture, then inspected and rendered)
```python
from pptx import Presentation
prs = Presentation("in.pptx")

# Hide slide 2 (1-based numbering for people, 0-based in code)
prs.slides[1]._element.set("show", "0")

# Move the last slide to the front
lst = prs.slides._sldIdLst
items = list(lst)
lst.remove(items[-1]); lst.insert(0, items[-1])

# Change chart numbers; the embedded workbook is rewritten too (checked with openpyxl)
from pptx.chart.data import CategoryChartData
chart = next(sh.chart for sh in prs.slides[2].shapes if sh.has_chart)
data = CategoryChartData(); data.categories = ["Q3", "Q4"]; data.add_series("Sales", (5.0, 7.0))
chart.replace_data(data)

prs.save("out.pptx")       # never the input path
```
Duplicate a slide (opened and rendered in the test; the copy is in no section until you add
its id; skip comments or two slides share one comment part):
```python
import copy
src = prs.slides[0]
new = prs.slides.add_slide(src.slide_layout)
for shp in list(new.shapes):
    shp._element.getparent().remove(shp._element)
for shp in src.shapes:
    new.shapes._spTree.append(copy.deepcopy(shp._element))
for rel in src.part.rels.values():
    if rel.reltype.endswith(("/notesSlide", "/slideLayout", "/comments")):
        continue
    if rel.is_external:
        new.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)  # new rId: fix r:id in copied XML
    else:
        new.part.rels._rels[rel.rId] = rel   # same rId, so copied r:embed/r:id still resolve
```
The external branch gets a new rId; in the test the only external link was a hyperlink whose
rId happened to differ, so check `inspect_pptx.py` output for broken actions after duplicating.
Private attributes (`_element`, `_sldIdLst`, `_rels`) are not public API and may change in a
later python-pptx release.

## Wrong vs right
- Wrong: `for shape in slide.shapes: shape.text = shape.text.replace(a, b)` (drops formatting,
  misses split runs, skips tables, groups and notes). Right: `pptx_replace.py --dry-run`, then `--expect`.
- Wrong: placing a text box at fixed coordinates for a title. Right: use the layout's title
  placeholder so the master's style and position apply and the outline shows the title.
- Wrong: `prs.save(path)` onto the user's file. Right: a new path, then inspect and render.
