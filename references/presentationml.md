# PresentationML internals

Load before hand-editing XML, or when a library result looks wrong. Facts marked *tested* were
checked on files written by python-pptx 1.0.2 and LibreOffice 25.2.3.2; the rest follow
ECMA-376 and [MS-PPTX] and are *documented, not tested here*.

## Package map
| Part | Holds | Points to |
| :--- | :--- | :--- |
| `[Content_Types].xml` | Content type per part; the main part type decides pptx/pptm/ppsx/potx | |
| `_rels/.rels` | Package relationships | `ppt/presentation.xml`, `docProps/*`, thumbnail, custom UI |
| `ppt/presentation.xml` | `p:sldMasterIdLst`, `p:notesMasterIdLst`, `p:sldIdLst` (show order), `p:sldSz`, `p:custShowLst`, section list in `p:extLst` | masters, slides, notes master, `presProps`, `viewProps`, theme, `vbaProject.bin`, comment authors |
| `ppt/slideMasters/slideMasterN.xml` | Master shapes, `p:txStyles` (title, body, other), `p:sldLayoutIdLst` | layouts, theme |
| `ppt/slideLayouts/slideLayoutN.xml` | Layout placeholders, `cSld@name` ("Title and Content") | its master |
| `ppt/slides/slideN.xml` | Slide shapes in `p:cSld/p:spTree`, `p:transition`, `p:timing` | its layout, notes slide, charts, pictures, media, OLE, comments, hyperlinks |
| `ppt/notesSlides/notesSlideN.xml` | Speaker notes: the `body` placeholder | its slide and the notes master |
| `ppt/charts/chartN.xml` | DrawingML chart with cached values | `../embeddings/*.xlsx` via `c:externalData` |
| `ppt/comments/*.xml` | Legacy `commentN.xml` (`p:cmLst`) or modern `modernComment_*.xml` (2018 schema) | |
| `ppt/media/`, `ppt/embeddings/` | Pictures, video, audio; embedded workbooks and OLE payloads | |
| `docProps/core.xml`, `app.xml`, `custom.xml` | Author, last modified by, company, slide titles list, custom properties | |

Content types of the main part (*tested* in the scripts' detection): presentation
`...presentationml.presentation.main+xml`, slide show `...slideshow.main+xml`, template
`...template.main+xml`; macro-enabled `application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml`
(and `.slideshow.`, `.template.` variants). PowerPoint expects the extension to match the type;
`pptx_sanitize.py --remove-macros` therefore requires a `.pptx` output.

## Slide list, ids and relationships
```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId7"/>   <!-- first slide shown -->
  <p:sldId id="257" r:id="rId9"/>
</p:sldIdLst>
```
- Order of `p:sldId` is the show order (*tested*: swapping two entries reorders `inspect_pptx.py`).
- `id` is a unique integer (ECMA-376 range 256 to 2147483647); it is what sections, custom shows
  and some add-ins refer to. Never reuse an id. Master ids live in a separate range above 2147483647.
- `r:id` names a relationship in `ppt/_rels/presentation.xml.rels` that targets the slide part.
- Removing a slide fully means: the `p:sldId`, the presentation relationship, the slide part and
  its `.rels`, its `Override` in `[Content_Types].xml`, its notes slide, section entries
  (`p14:sldId`), custom show entries (`p:custShowLst//p:sld`), and links from other slides
  (`ppaction://hlinksldjump` relationships of type `.../slide`). *Tested*: python-pptx's
  common delete recipe left `slide2.xml` in the package because slide 1 still linked to it,
  and left a stale section id. `inspect_pptx.py` reports `unlisted_slide_parts` and
  `stale_section_ids`; `pptx_sanitize.py --remove-hidden-slides` removes both, and
  `pptx_slides.py --delete` does the whole list above in one step.

## Masters, layouts, placeholders and inheritance
- Each slide has exactly one layout relationship; each layout one master.
- A placeholder is a `p:sp` whose `p:nvSpPr/p:nvPr/p:ph` has `type` (title, ctrTitle, subTitle,
  body, pic, chart, tbl, dt, ftr, sldNum, ...; default body) and `idx`.
- A slide placeholder inherits position (`a:xfrm`), geometry, text body properties and list
  styles from the layout placeholder that matches it (by `idx`, then by `type`), which inherits
  from the master; text formatting falls back to `p:txStyles` on the master.
- A slide placeholder with its own `a:xfrm` was moved or resized on that slide; it no longer
  follows layout changes. `inspect_pptx.py` reports `own_position`.
- An **orphan** placeholder (no match on layout or master, often after switching layout) keeps
  its text but loses inherited position and style; it can render at 0,0 or with default fonts.
  `inspect_pptx.py` reports `inherits_from: "none"` (*tested* with a fixture).
- Footer, date and slide number placeholders show only if the slide has them; turning them on
  for all slides is a per-slide change, not a master change (documented, not tested here).

## Text model (DrawingML)
```xml
<p:txBody><a:bodyPr/><a:lstStyle/>
  <a:p>
    <a:pPr lvl="1"/>                                   <!-- bullet level -->
    <a:r><a:rPr lang="en-US" b="1"/><a:t>Q3 rev</a:t></a:r>
    <a:r><a:rPr lang="en-US"/><a:t>enue</a:t></a:r>
    <a:br/>                                             <!-- soft line break -->
    <a:fld id="{...}" type="slidenum"><a:t>4</a:t></a:fld>
    <a:endParaRPr lang="en-US"/>
  </a:p>
</p:txBody>
```
- Runs split anywhere; matching must join runs inside a paragraph (`pptx_replace.py`).
- `a:fld` is a field: slide number, date. Its `a:t` is a cached result; do not edit it.
- python-pptx returns `a:br` as `\x0b` (vertical tab) and includes field text in `.text` (*tested*).
- Autofit (`a:normAutofit fontScale=...`) is computed by PowerPoint when it renders; a library
  edit that adds text does not shrink it, and LibreOffice may fit differently. Render and look.

## Notes, comments, sections, hidden slides
- **Notes**: slide relationship of type `.../notesSlide`. The text is the notes slide's `body`
  placeholder. A notes slide also contains a slide image placeholder and may hold header/footer
  text. The notes master (`ppt/notesMasters/`) is shared.
- **Comments**: legacy, slide relationship `.../relationships/comments` to `p:cmLst`, authors in
  `ppt/commentAuthors.xml`. Modern (Microsoft 365), relationship
  `http://schemas.microsoft.com/office/2018/10/relationships/comments`, content type
  `application/vnd.ms-powerpoint.comments+xml`, authors in `ppt/authors.xml` (from [MS-PPTX];
  only the legacy form is in the test fixture).
- **Sections**: `p:extLst/p:ext[@uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"]/p14:sectionLst`,
  each `p14:section@name` with `p14:sldIdLst/p14:sldId@id` matching `p:sldId@id`. A new slide
  is in no section until you add its id.
- **Hidden slide**: `<p:sld show="0">`. Hidden slides stay in the file with all their text,
  notes and comments. LibreOffice 25.2.3.2 left them out of the PDF (*tested*); PowerPoint's PDF
  export has an "Include hidden slides" option (documented, not tested here).

## Actions and links
`a:hlinkClick` (on a run or on `p:cNvPr`) and `a:hlinkHover` carry an optional `r:id` and `action`:
| `action` | Effect |
| :--- | :--- |
| none, with external `r:id` | Opens the URL or file. Scheme matters: `file:`, UNC `\\host\share`, `javascript:` are risky |
| `ppaction://hlinksldjump` | Jump to the slide named by an internal `.../slide` relationship |
| `ppaction://hlinkshowjump?jump=nextslide` (and `previousslide`, `firstslide`, `lastslide`, `endshow`) | Navigation |
| `ppaction://program` | Runs the program the relationship targets |
| `ppaction://macro?name=X` | Runs a VBA macro |
| `ppaction://ole?verb=0` | Activates an embedded object |
| `ppaction://customshow?id=0` | Starts a custom show |
`inspect_pptx.py` flags program, macro, ole and non-http(s)/mailto schemes (*tested* with a fixture).

## Charts, media, OLE
See `charts-media.md`. In short: a chart is a `p:graphicFrame` whose graphic data points to a
chart part; the chart part points to an embedded workbook. Video and audio are `p:pic` with
`a:videoFile`/`a:audioFile` (`r:link`) plus a `p14:media` extension (`r:embed`). OLE objects
are `p:graphicFrame/.../p:oleObj`, usually wrapped in `mc:AlternateContent` with a fallback picture.

## Repair prompts ("PowerPoint found a problem with content")
Usual causes, in the order to check: a relationship target that does not exist; a part without
a content type; a `p:sldId` id reused or below 256; a shape `id` (`p:cNvPr@id`) duplicated on
one slide; a section listing an id that is not in `p:sldIdLst`; a macro part in a `.pptx`. The
smoke test helper `consistent()` checks the first two and the section ids.
