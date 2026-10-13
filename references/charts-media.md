# Charts, media, animations, transitions, sections

## Charts and their embedded workbooks
A chart on a slide is three things (*tested* with python-pptx 1.0.2 output):
1. `p:graphicFrame` on the slide, graphic data `uri=".../drawingml/2006/chart"`, `c:chart r:id`.
2. `ppt/charts/chartN.xml`: series formulas (`c:f`, for example `Sheet1!$B$2:$B$3`) and
   **cached values** (`c:numCache`, `c:strCache`). Renderers draw the cache.
3. `ppt/embeddings/Microsoft_Excel_Sheet*.xlsx`, linked by `c:externalData r:id` with a
   relationship of type `.../relationships/package`. PowerPoint opens it on "Edit Data".

Consequences:
- Editing the cache only: the slide looks right, but "Edit Data" in PowerPoint shows the old
  numbers and the next edit there can bring them back. Editing the workbook only: nothing
  changes on the slide until PowerPoint refreshes. Change both: `chart.replace_data()` does
  (*tested*: the embedded workbook held the new values, checked with openpyxl 3.1.5).
- The workbook can hold more than the chart shows (other sheets, hidden rows, the full
  dataset someone pasted). Before sharing, open it (it is a normal .xlsx in the zip) or drop it
  with `pptx_sanitize.py --remove-chart-workbooks`: the chart keeps its cached look, "Edit
  Data" stops working.
- A chart can link to an **external** workbook instead (relationship `TargetMode="External"`).
  PowerPoint may offer to update links on open (documented, not tested). `inspect_pptx.py`
  lists it under `external`; `--remove-active` drops the link and keeps the cache.
- Text replace does not change chart text; `pptx_replace.py` reports matches there under
  `not_replaced.charts`. Edit category names through the chart data.

## Media (video, audio)
python-pptx `add_movie` wrote (*tested*): a `p:pic` with `a:videoFile r:link="rId3"` (type
`.../2006/relationships/video`) and an extension `p14:media r:embed="rId2"` (type
`http://schemas.microsoft.com/office/2007/relationships/media`), both to `ppt/media/media1.mp4`,
plus a poster image. A **linked** video uses `TargetMode="External"` with a path or URL; it
plays only where that path resolves and leaks the path. `inspect_pptx.py` reports media with
`linked: true|false`. Trimming, fades and bookmarks live in `p14:media` extensions
(documented, not tested). LibreOffice renders the poster frame, not the video.

## Animations and transitions
- Transition: `p:transition` on the slide (`spd`, `advTm` for auto-advance in ms, a child such
  as `p:fade`), often inside `mc:AlternateContent` with a `p14`/`p15` choice (morph is a p159
  extension in Microsoft 365, documented, not tested). `inspect_pptx.py` reports the type and
  a transition sound.
- Animations: `p:timing` tree of `p:par`/`p:cTn` time nodes; `presetClass` (entr, exit, emph,
  path) marks an effect, and `p:bldLst` lists builds. Effects target shapes by `spid`
  (`p:cNvPr@id`). Deleting or replacing a shape leaves a dangling target that PowerPoint repairs
  or drops; keep shape ids when editing. `inspect_pptx.py` counts effects.
- A PDF or PNG shows the final state of a slide, not animation steps, and no transition.

## Sections
Stored in `presentation.xml` (see `presentationml.md`). Rules:
- Every `p14:sldId@id` must exist in `p:sldIdLst`; `inspect_pptx.py` reports `stale_section_ids`.
- Slides added by python-pptx are in no section. Append `<p14:sldId id="..."/>` to the section
  you want, using the new `p:sldId@id`.
- Section names are visible to anyone who opens the deck; they can leak ("Do not show: pricing").

## Hidden slides
`show="0"` on `p:sld`. They are in the file, can be shown with one click, carry notes and
comments, and are listed by text extractors such as markitdown without a marker (see
`extraction.md`). Ask the user whether hidden slides should be shared, removed, or unhidden.
