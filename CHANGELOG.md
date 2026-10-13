# Changelog

## 0.1.2 (2026-10-13)
Fixes from review round 2. Not published.
- Links (`pptlinks.py`): e-mail addresses are parsed before domains, so `jane.doe@corp.com` over
  its own mailto no longer fails the share check (and is no longer removed), and
  `security@paypal.com` over `mailto:x@evil.io` is flagged. Blank-looking characters (U+034F,
  Hangul fillers, Khmer U+17B4/U+17B5) count as invisible. Display domains may be non-ASCII and
  are compared in IDNA form; a label that mixes scripts (`pаypal.com`) is `mixed_script`.
  `paypal.com@evil.io` stays `userinfo`; a person's dotted address does not.
- Every writer refuses an OUT that is the input (realpath and samefile) or a symlink, even with
  `--force`. `pptx_build.py` could overwrite its template that way; `pptx_render.py --to pptx`
  could overwrite its input.
- `pptx_build.py`: image paths must stay in the spec folder (`--allow-image-outside` to lift);
  bad images, bad `level` values and python-pptx errors give exit 2 with JSON, not a traceback;
  the .potx content type is changed through the parsed `[Content_Types].xml`.
- `pptx_slides.py` JSON adds `final_order` (input numbers in the new order). `pptx_replace.py`
  matches add `after`, so a dry run previews each change.
- Docs: share procedure with stop conditions and a "common jobs" table in recipes.md; SKILL.md
  says what `--check` does and does not mean, the encrypted-deck answer, build stop rules and the
  memory-cap retry; automation.md marked as user-side; notes loop exits on the first body
  placeholder with text; security.md carries the suffix-list caveat and the new link reasons.
- The unsandboxed LibreOffice test is named as a positive control, never to copy.
- Self-check: `tests/smoke_test.py` and the fixture ship in the package; skill.json gains
  `keywords` and `selfcheck`.
- Tests: 53 (was 45).

## 0.1.1 (2026-10-13)
Fixes from review round 1. Not published.
- Security: VBA found by relationship and content type, not file name; a macro-enabled content
  type is its own risk. Click and hover actions and link schemes (javascript:, file:, data:,
  UNC, relative) are checked in every XML part, including masters, layouts and notes. External
  targets are collected from every .rels file (theme, SmartArt, presProps included). New link
  text check: shown domain differs from target, userinfo tricks, invisible characters
  (`scripts/pptlinks.py`, written for this skill). `--remove-active` removes all of these.
- `pptx_replace.py`: refuses control, bidi and zero-width characters in replacement text; a bad
  pattern or child error is reported as such, not as a timeout.
- Metadata: `--check-clean` and `--strip-metadata` use one field list, now including title,
  revision, Template, TotalTime, language, version.
- `--check-clean` and sanitize results carry `clean_scope` (`checked`, `not_checked`) and say
  that ok is not a safety verdict.
- `pptx_render.py`: input checked by content, not extension (zip caps for .odp, renamed files
  refused), `--memory-mb` address-space cap (default 4096), `--to pptx` result must be
  inspected and sanitized before use.
- Every writer refuses an output extension that does not match the package type.
- New `pptx_slides.py`: delete (complete, incl. notes, links, sections, custom shows), hide,
  unhide, move, `--expect-slides`, consistency check before writing.
- New `pptx_build.py`: build from the user's .pptx or .potx by placeholder idx, `--list-layouts`.
- SKILL.md: shorter description, quick start with real output, platform matrix, read/write
  marks, input and output per script, no-render path, common mistakes, reference load table.
- Tests: 45 (was 29), including doc checks that run the duplicate-slide recipe; tests use a temp copy of the fixture and check the tracked file is
  unchanged; `make_fixtures.py` writes to a temp folder unless `--write`.
- `requirements.txt` pins `lxml>=5`; `skill.json` has `tested_with` and `not_tested`.

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
