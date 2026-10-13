# Automating PowerPoint itself: VBA, Office.js, Microsoft Graph

**Documented, not executed.** Nothing here was run against PowerPoint or a Microsoft 365
tenant. Office.js member names and requirement sets were checked against `@types/office-js`
1.0.611 (npm, 2026-10-13); Graph against Microsoft Learn on 2026-10-13; VBA against
Microsoft Learn. Verify on the user's build before relying on any of it.

## Which surface
| Situation | Use |
| :--- | :--- |
| File on disk, no PowerPoint | This skill's scripts and python-pptx |
| PowerPoint desktop on Windows or Mac, user present | VBA (or COM from PowerShell/Python on Windows) |
| Add-in that must work in PowerPoint on the web, Windows, Mac, iPad | Office.js PowerPoint API |
| File in OneDrive or SharePoint, server side | Microsoft Graph (download, upload, convert to PDF) |

## VBA (PowerPoint object model)
```vb
' Export PDF (ppFixedFormatTypePDF = 2). PrintHiddenSlides defaults to msoFalse; the docs mark
' PrintRange as required but allow Nothing; IncludeDocProperties defaults to False.
ActivePresentation.ExportAsFixedFormat "C:\out\deck.pdf", ppFixedFormatTypePDF

' Slide images
ActivePresentation.Slides(1).Export "C:\out\slide1.png", "PNG", 1920, 1080

' Hide slide 3; read notes of slide 1
ActivePresentation.Slides(3).SlideShowTransition.Hidden = msoTrue
Debug.Print ActivePresentation.Slides(1).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text

' Replace text in every shape that has text (runs are handled by PowerPoint)
Dim s As Slide, sh As Shape
For Each s In ActivePresentation.Slides
  For Each sh In s.Shapes
    If sh.HasTextFrame Then sh.TextFrame.TextRange.Replace FindWhat:="Q3", ReplaceWhat:="Q4"
  Next
Next

' Remove comments, personal info and more before sharing
ActivePresentation.RemoveDocumentInformation ppRDIAll
```
Notes: `TextRange.Replace` replaces the first match per call; loop until it returns Nothing for
all matches. Group shapes and tables need their own loops (`GroupItems`, `Table.Cell`). VBA
runs only in desktop PowerPoint, not on the web. A macro saved into a deck makes it `.pptm`.

## Office.js PowerPoint API (requirement set per member, from @types/office-js 1.0.611)
| Member | Set |
| :--- | :--- |
| `Office.context.document.getFileAsync` (whole file) | Common API |
| `PowerPoint.createPresentation(base64)` | PowerPointApi 1.1 |
| `presentation.insertSlidesFromBase64(...)` | PowerPointApi 1.2 |
| `slide.delete()` | PowerPointApi 1.2 |
| `slides.add(options)`, `slideMasters`, `layouts`, `tags`, `shape.delete()` | PowerPointApi 1.3 |
| `shapes.addTextBox`, `addGeometricShape`, `addLine`, `textFrame.textRange` | PowerPointApi 1.4 |
| `getSelectedSlides`, `setSelectedSlides` | PowerPointApi 1.5 |
| `Hyperlink` | PowerPointApi 1.6 |
| `customXmlParts` | PowerPointApi 1.7 |
| `slide.exportAsBase64()`, `slide.getImageAsBase64()`, `shapes.addTable`, `slide.applyLayout`, `slide.moveTo`, `placeholderFormat`, bindings | PowerPointApi 1.8 |
| `slides.exportAsBase64Presentation(...)`, `shape.getImageAsBase64()`, `textRange.hyperlinks` | PowerPointApi 1.10 |
Not in 1.0.611 for PowerPoint: speaker notes, hidden-slide flag, sections, comments,
animations, transitions. Use VBA or edit the file.
```ts
await PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const shapes = slides.items[0].shapes;
  shapes.load("items/name,items/type");
  await context.sync();
  for (const s of shapes.items) console.log(s.name, s.type);
});
```
Check support at run time: `Office.context.requirements.isSetSupported("PowerPointApi", "1.8")`.

## Microsoft Graph
- Download: `GET /drives/{drive-id}/items/{item-id}/content`. Upload (small):
  `PUT /drives/{drive-id}/items/{parent-id}:/{name}.pptx:/content`.
- Convert to PDF: `GET /drives/{drive-id}/items/{item-id}/content?format=pdf`; returns `302`
  to a short-lived URL. Listed sources include pptx, ppsx, ppt, pps, odp (not pptm or potx).
  **Size limit for format conversion: 2 MB.** `format=jpg` with `width` and `height` gives an
  image of the file.
- Least privilege: delegated `Files.Read` to read or convert; `Files.ReadWrite` only to write.
  Never log tokens or the pre-authenticated download URL.
- Graph renders with Microsoft's service, so it is closer to PowerPoint than LibreOffice, but
  the file leaves the machine: ask first.
