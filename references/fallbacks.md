# Fallbacks

| Missing or different | Do this |
| :--- | :--- |
| lxml | `pip install "lxml>=5"` (ask first in a managed environment). Without it no script runs; read the deck with python-pptx if present, but do not sanitize by hand |
| python-pptx | Not needed for inspect, replace, sanitize, slides, render. Needed for `pptx_build.py` and custom shapes |
| LibreOffice | No rendering or legacy conversion. Ask the user to export PDF from PowerPoint, or install `libreoffice-impress` (Debian/Ubuntu). Do not install a full desktop suite without asking |
| bubblewrap, or bwrap cannot create namespaces (some containers disable user namespaces) | `pptx_render.py` refuses (exit 2). Options: install bubblewrap; run in a VM or container that allows it; or have the user export from PowerPoint. Never run plain `soffice` on a deck you did not create |
| pdftoppm | PDF still works; PNG needs `poppler-utils` |
| macOS | inspect, replace, sanitize, slides work (pure Python, not tested on macOS). Render: PowerPoint File > Export > PDF, or `ExportAsFixedFormat` / `Slide.Export` in VBA (`automation.md`). This skill has no unsandboxed render mode, on purpose. If no export is possible, follow the no-render path in SKILL.md and say layout was not checked |
| Windows | Same as macOS; COM automation from PowerShell is possible on a machine with PowerPoint (documented, not tested) |
| Legacy `.ppt`, `.pps`, `.pot` | Scripts refuse with exit 2. Convert a copy: `pptx_render.py old.ppt old.pptx --to pptx` (Linux, sandboxed), then inspect and, before sharing, sanitize the result: LibreOffice can carry macros, links and notes across. LibreOffice re-saves everything; keep the original |
| `.odp` (Impress) | Same as `.ppt`: convert, then inspect |
| `.potx` template | Inspect and sanitize directly. `pptx_build.py` reads it through a copy with the presentation content type (tested: python-pptx 1.0.2 refuses the template type with ValueError) |
| `.ppsx` slide show | Treat as a deck. It opens as a show in PowerPoint |
| Encrypted (password) | Exit 2 with "password-protected". Ask for an unprotected copy; do not try to crack it. msoffcrypto-tool can decrypt with the password (not tested here) |
| PowerPoint for the web only | No VBA. Office.js add-ins work; or download, edit with this skill, upload a new version (keep the original) |
| Files in OneDrive or SharePoint | Graph download, process locally, upload under a new name; `?format=pdf` for a Microsoft-rendered PDF up to 2 MB (see `automation.md`) |
| Google Slides or Keynote | Export to .pptx first; expect layout differences |

## No-render path
When `pptx_render.py --check` is not ok: finish the edit, inspect the output, then ask the user to open the copy
(or export a PDF from PowerPoint, or use Graph `?format=pdf` for OneDrive files under 2 MB). In the report say
"layout not checked" and name the slides whose text changed length the most. Never fall back to plain `soffice`.
