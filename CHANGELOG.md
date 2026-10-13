# Changelog

## 0.1.0 (2026-10-13)
First release, not yet published to npm or GitHub.
- SKILL.md router: decision guide, safe workflow, never-do and ask-first lists, exit code table,
  verification checklist, macOS and Windows note.
- References: PresentationML internals, python-pptx limits with tested workarounds, charts and
  media, sandboxed rendering, security, four worked examples, extraction comparison, VBA /
  Office.js / Graph automation, fallbacks, troubleshooting.
- Scripts (Python, lxml only): `inspect_pptx.py`, `pptx_replace.py` (split runs, `--expect`,
  regex timeout), `pptx_sanitize.py` (categories, `--all`, `--check-clean`, residue exit 1),
  `pptx_render.py` (LibreOffice in bubblewrap: PDF, PNG, notes pages, legacy conversion).
- Tests: `tests/smoke_test.py` with `tests/fixtures/deck.pptx`, including canaries that show an
  unsandboxed LibreOffice run fetching a linked http picture and embedding a linked local file,
  and that the sandboxed run does neither.
