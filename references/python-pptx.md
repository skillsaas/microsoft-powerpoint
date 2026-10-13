# python-pptx: what it does and does not do

Tested with python-pptx 1.0.2 on Python 3.13 unless marked otherwise. Load when writing
python-pptx code beyond `pptx_build.py`, or when its output surprises you.

First choice for a new deck: `pptx_build.py --list-layouts TEMPLATE`, then a JSON spec that fills
placeholders by idx (see SKILL.md). Write code only for what the spec cannot express (tables,
charts, shapes).

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
| Delete a slide (`prs.slides` has only `add_slide`, `get`, `index`) | Save, then `pptx_slides.py IN OUT --delete N` (removes sldId, rel, notes, links to it, section and custom show entries). The common `drop_rel` + remove `sldId` recipe leaves the part in the file when another slide links to it, and leaves stale section ids (tested) |
| Reorder slides | `pptx_slides.py IN OUT --move FROM:TO` (keeps sections in order); in code, move `p:sldId` elements in `prs.slides._sldIdLst` (below) |
| Hide or unhide a slide (no `hidden` property) | `pptx_slides.py --hide N` / `--unhide N`; in code `slide._element.set("show", "0")`, delete the attribute to unhide |
| Duplicate a slide | Deep-copy shapes, relate the new slide to the same targets, and rewrite every `r:` attribute to the new rIds (below) |
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
Duplicate a slide. Tested on a slide with a picture, an internal slide jump and an external
hyperlink: every rId in the copy was remapped, inspect showed no broken references, and the
picture rendered. The copy is in no section and has no notes or comments.
```python
import copy
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def duplicate(prs, index):
    src = prs.slides[index]
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):                      # drop the layout's empty placeholders
        shp._element.getparent().remove(shp._element)
    remap = {}
    for rel in src.part.rels.values():
        if rel.reltype.endswith(("/notesSlide", "/slideLayout", "/comments")):
            continue
        if rel.is_external:
            remap[rel.rId] = new.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            remap[rel.rId] = new.part.relate_to(rel.target_part, rel.reltype)
    for shp in src.shapes:
        el = copy.deepcopy(shp._element)
        for node in el.iter():                        # r:embed, r:link, r:id ...
            for key, val in list(node.attrib.items()):
                if key.startswith("{%s}" % R) and val in remap:
                    node.set(key, remap[val])
        new.shapes._spTree.append(el)
    return new                                        # added at the end; move with pptx_slides.py
```
Then run `inspect_pptx.py OUT --json` and check `risks` for `broken-slide-reference`.
Private attributes (`_element`, `_sldIdLst`, `_rels`) are not public API and may change in a
later python-pptx release.

## Wrong vs right
- Wrong: `for shape in slide.shapes: shape.text = shape.text.replace(a, b)` (drops formatting,
  misses split runs, skips tables, groups and notes). Right: `pptx_replace.py --dry-run`, then `--expect`.
- Wrong: placing a text box at fixed coordinates for a title. Right: use the layout's title
  placeholder so the master's style and position apply and the outline shows the title.
- Wrong: `prs.save(path)` onto the user's file. Right: a new path, then inspect and render.
