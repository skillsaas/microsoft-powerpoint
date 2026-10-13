# Automating PowerPoint itself: VBA, Office.js, Microsoft Graph

**For the user's side.** These run inside PowerPoint, an add-in or a Microsoft 365 tenant, not in
the agent's shell. Give them to the user to run (or run them only where the user has set that
up and asked), and say they are documented, not tested here.

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
Write next to the saved deck, never to a hard-coded drive. An unsaved deck has an empty `Path`.
```vb
Sub ExportDeck()
  If ActivePresentation.Path = "" Then Err.Raise 5, , "Save the presentation first"
  Dim sep As String: sep = Application.PathSeparator   ' "\" on Windows, "/" on Mac
  Dim base As String: base = ActivePresentation.Path & sep & "export"
  ' PDF without hidden slides. Named arguments, so no positional guessing.
  ActivePresentation.ExportAsFixedFormat Path:=base & ".pdf", _
      FixedFormatType:=ppFixedFormatTypePDF, PrintHiddenSlides:=msoFalse
  ' One slide as an image
  ActivePresentation.Slides(1).Export base & "-slide1.png", "PNG", 1920, 1080
End Sub
```
Mac: PowerPoint is sandboxed, so writing outside the deck's folder can prompt for access or
fail; keep outputs beside the deck (documented behaviour of sandboxed Mac apps, not tested here).

```vb
' Speaker notes of slide 1: find the body placeholder by type, not by index.
Function NotesText(sl As Slide) As String
  Dim sh As Shape
  For Each sh In sl.NotesPage.Shapes.Placeholders
    If sh.PlaceholderFormat.Type = ppPlaceholderBody And sh.HasTextFrame Then
      NotesText = sh.TextFrame.TextRange.Text
      Exit For
    End If
  Next
End Function

' Hide slide 3
ActivePresentation.Slides(3).SlideShowTransition.Hidden = msoTrue
```

Replace everywhere. `TextRange.Replace` changes the first match after `After` and returns the
replaced range, or Nothing; loop per text range. Groups and tables hold their own text:
```vb
Sub ReplaceAll(findWhat As String, replaceWhat As String)
  Dim sl As Slide, sh As Shape
  For Each sl In ActivePresentation.Slides
    For Each sh In sl.Shapes
      ReplaceInShape sh, findWhat, replaceWhat
    Next
  Next
End Sub

Sub ReplaceInShape(sh As Shape, f As String, r As String)
  Dim i As Long, j As Long
  If sh.Type = msoGroup Then
    For i = 1 To sh.GroupItems.Count: ReplaceInShape sh.GroupItems(i), f, r: Next
  ElseIf sh.HasTable Then
    For i = 1 To sh.Table.Rows.Count
      For j = 1 To sh.Table.Columns.Count
        ReplaceInRange sh.Table.Cell(i, j).Shape.TextFrame.TextRange, f, r
      Next
    Next
  ElseIf sh.HasTextFrame Then
    ReplaceInRange sh.TextFrame.TextRange, f, r
  End If
End Sub

Sub ReplaceInRange(tr As TextRange, f As String, r As String)
  Dim hit As TextRange, pos As Long
  Do
    Set hit = tr.Replace(FindWhat:=f, ReplaceWhat:=r, After:=pos)
    If hit Is Nothing Then Exit Do
    pos = hit.Start + hit.Length - 1   ' continue after the replacement, so "a" -> "aa" ends
  Loop
End Sub
```
Notes pages, masters and charts are not covered by that loop; say so in the report.

Before sharing, `RemoveDocumentInformation` changes the open deck in place (comments, notes,
personal info, depending on the type). Ask first, work on a saved copy (`SaveCopyAs`, then open
the copy), and never run it on the user's only file:
```vb
ActivePresentation.RemoveDocumentInformation ppRDIAll   ' on the copy, after the user agreed
```
VBA runs only in desktop PowerPoint, not on the web. A macro saved into a deck makes it `.pptm`.

## Office.js PowerPoint API (requirement set per member, from @types/office-js 1.0.611)
| Member | Set |
| :--- | :--- |
| `Office.context.document.getFileAsync` (whole file) | Common API |
| `PowerPoint.createPresentation(base64)` | PowerPointApi 1.1 |
| `presentation.insertSlidesFromBase64(base64, options)` | PowerPointApi 1.2 |
| `slide.delete()` | PowerPointApi 1.2 |
| `slides.add(options)`, `slideMasters`, `layouts`, `tags`, `shape.delete()` | PowerPointApi 1.3 |
| `shapes.addTextBox`, `addGeometricShape`, `addLine`, `textFrame.textRange` | PowerPointApi 1.4 |
| `getSelectedSlides`, `setSelectedSlides` | PowerPointApi 1.5 |
| `Hyperlink` | PowerPointApi 1.6 |
| `customXmlParts` | PowerPointApi 1.7 |
| `slide.exportAsBase64()`, `slide.getImageAsBase64()`, `shapes.addTable`, `slide.applyLayout`, `slide.moveTo`, `placeholderFormat`, bindings | PowerPointApi 1.8 |
| `TableRow.delete()`, `TableColumn.delete()` and other table edits | PowerPointApi 1.9 |
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

`insertSlidesFromBase64` details (Microsoft Learn "Insert slides from another PowerPoint
presentation" and @types/office-js 1.0.611):
- The argument is raw base64 of a .pptx. `FileReader.readAsDataURL` gives a `data:...;base64,`
  prefix; strip everything up to and including `base64,`.
- `targetSlideId` (insert after it; default: at the start) and each `sourceSlideIds` entry take
  one of `nnn#`, `#mmmmmmmmm` or `nnn#mmmmmmmmm` (slide id, creation id, or both), for
  example `267#`. A missing id throws `SlideNotFound` and nothing is inserted.
- Slides are inserted in their order in the source deck, whatever the array order.
- `formatting`: `"KeepSourceFormatting"` (default) or `"UseDestinationTheme"`.
- The whole source file sits in memory as a string about 4/3 of its size; keep sources small.
- The matching export is `slides.exportAsBase64Presentation(ids)` (PowerPointApi 1.10).


## Microsoft Graph
- Download: `GET /drives/{drive-id}/items/{item-id}/content`. Upload (small):
  `PUT /drives/{drive-id}/items/{parent-id}:/{name}.pptx:/content`.
- Convert to PDF: `GET /drives/{drive-id}/items/{item-id}/content?format=pdf`. The answer is a
  `302` to a short-lived, pre-authenticated URL. Follow it at once and do not send the
  `Authorization` header to that host:
  ```bash
  curl -sS -L -o deck.pdf -H "Authorization: Bearer $TOKEN" \
    "https://graph.microsoft.com/v1.0/drives/$DRIVE/items/$ITEM/content?format=pdf"
  ```
  curl drops a custom `Authorization` header when `-L` moves to another host (curl 7.58 and
  later; older versions leak it). With another client, read `Location` and fetch it without
  the header. Listed sources include pptx, ppsx, ppt, pps, odp (not pptm or potx).
- **Size limit for format conversion: 2 MB.** Bigger files: download and render locally with
  `pptx_render.py` (Linux, sandboxed), or ask the user to export from PowerPoint. Ask before
  either: one sends nothing out, the other needs their PowerPoint.
- `format=jpg` with `width` and `height` is documented to give an image of the file (not tested here).
- Least privilege: delegated `Files.Read` to read or convert; `Files.ReadWrite` only to write.
  Never log tokens or the pre-authenticated download URL.
- Graph renders with Microsoft's service, closer to PowerPoint than LibreOffice, but the file
  leaves the machine: ask first.
