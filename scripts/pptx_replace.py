#!/usr/bin/env python3
"""Find and replace slide text that PowerPoint split across runs, keeping run formatting.

Usage:
  python pptx_replace.py IN OUT --find TEXT --replace TEXT [--regex] [--ignore-case]
         [--expect N] [--notes] [--masters] [--slides 1,3-5] [--dry-run] [--force]
         [--timeout S] [--json]

Why: PowerPoint stores "Q3 revenue" as several a:r runs whenever spelling marks, autocorrect,
language tags or a formatting change touched the words, so python-pptx run-level replace and a
plain XML string replace both miss it, and text_frame.text = "..." loses run formatting
(tested with python-pptx 1.0.2: a bold run comes back with bold None).

How: inside each paragraph (a:p) the text of consecutive runs is joined, matched, and the
replacement is written into the run where the match starts; the matched characters are cut
from the following runs, so each run keeps its own formatting. A match never crosses a line
break (a:br) or a field (a:fld: slide number, date); those are counted as "skipped".

Scope: all slides in show order, hidden slides included (reported). --notes adds speaker notes.
--masters adds slide layouts and masters (changes every slide that inherits that text).
Text in charts, SmartArt and embedded objects is never changed; matches there are counted
under "not_replaced" so you can tell the user.

Safety: writes only OUT (never the input; refuses if OUT exists unless --force, which keeps no
backup). --expect N refuses to write unless exactly N replacements would be made. --regex runs
in a child process killed after --timeout seconds (default 5), so a catastrophic pattern from
a user or a file cannot hang the agent. Parts that are not changed are copied byte for byte.

Exit codes: 0 written (or dry run done), 1 count differs from --expect or no match (nothing
written), 2 unsupported input, bad pattern, regex timeout or output problem, 3 safety refusal.
"""
import argparse
import multiprocessing as mp
import re
import sys

import pptpkg as pk
from pptpkg import NS, q

def _spans(pattern, flags, repl, is_regex, texts, out):
    rx = re.compile(pattern, flags)
    res = []
    for t in texts:
        res.append([(m.start(), m.end(), m.expand(repl) if is_regex else repl) for m in rx.finditer(t)
                    if m.end() > m.start()])
    out.put(res)


def find_all(pattern, flags, repl, is_regex, texts, timeout):
    ctx = mp.get_context("fork" if sys.platform.startswith("linux") else "spawn")
    qu = ctx.Queue()
    pr = ctx.Process(target=_spans, args=(pattern, flags, repl, is_regex, texts, qu), daemon=True)
    pr.start()
    try:
        res = qu.get(timeout=timeout)
    except Exception:
        pr.kill()
        raise pk.Unsupported("pattern did not finish within %s s (possible catastrophic regex); nothing written" % timeout)
    pr.join(1)
    if isinstance(res, Exception):
        raise res
    return res


def paragraph_chunks(p):
    """Split a paragraph into chunks of consecutive a:r runs; a:br and a:fld break chunks.

    Also returns the paragraph as a reader sees it (a:br as a space, field text included),
    used only to count matches that cross a break or a field and so are not replaced.
    """
    chunks, cur, full = [], [], []
    for child in p:
        if child.tag == q("a", "r"):
            t = child.find("a:t", NS)
            if t is not None:
                cur.append(t)
                full.append(t.text or "")
        elif child.tag in (q("a", "br"), q("a", "fld")):
            if cur:
                chunks.append(cur)
            cur = []
            ft = child.find("a:t", NS) if child.tag == q("a", "fld") else None
            full.append(" " if ft is None else (ft.text or ""))
    if cur:
        chunks.append(cur)
    return chunks, "".join(full)


def apply(ts, matches):
    """Rewrite the a:t nodes of one chunk for the given (start, end, replacement) matches."""
    bounds, pos = [], 0
    for t in ts:
        n = len(t.text or "")
        bounds.append((pos, pos + n))
        pos += n
    for start, end, repl in sorted(matches, reverse=True):
        first = True
        for i, (b0, b1) in enumerate(bounds):
            if b1 <= start or b0 >= end:
                continue
            s, e = max(start, b0) - b0, min(end, b1) - b0
            txt = ts[i].text or ""
            ts[i].text = txt[:s] + (repl if first else "") + txt[e:]
            first = False
        # Recompute bounds after edit
        bounds, pos = [], 0
        for t in ts:
            n = len(t.text or "")
            bounds.append((pos, pos + n))
            pos += n
    for t in ts:  # keep spaces at the run edges
        if t.text and (t.text != t.text.strip()):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def parse_slides(spec, n):
    if not spec:
        return None
    out = set()
    for bit in spec.split(","):
        if "-" in bit:
            a, b = bit.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(bit))
    bad = [i for i in out if i < 1 or i > n]
    if bad:
        raise pk.Unsupported("slide numbers out of range 1..%d: %s" % (n, bad))
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--find", required=True)
    ap.add_argument("--replace", required=True)
    ap.add_argument("--regex", action="store_true")
    ap.add_argument("--ignore-case", action="store_true")
    ap.add_argument("--expect", type=int)
    ap.add_argument("--notes", action="store_true")
    ap.add_argument("--masters", action="store_true")
    ap.add_argument("--slides")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--timeout", type=float, default=5.0)
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    try:
        if not a.find:
            raise pk.Unsupported("--find is empty")
        pattern = a.find if a.regex else re.escape(a.find)
        flags = re.IGNORECASE if a.ignore_case else 0
        try:
            re.compile(pattern, flags)
        except re.error as exc:
            raise pk.Unsupported("bad regex: %s" % exc)
        repl = a.replace
        pkg = pk.Package(a.input)
        slides = pkg.slides()
        pick = parse_slides(a.slides, len(slides))
        targets = []  # (part, label, hidden)
        for s in slides:
            if s["part"] and (pick is None or s["index"] in pick):
                targets.append((s["part"], "slide %d" % s["index"], s["hidden"]))
                if a.notes:
                    for r in pkg.rel_by_type(s["part"], "/notesSlide"):
                        if pkg.exists(r["part"]):
                            targets.append((r["part"], "notes of slide %d" % s["index"], s["hidden"]))
        if a.masters:
            for m in pkg.rel_by_type(pkg.main, "/slideMaster"):
                targets.append((m["part"], "master %s" % m["part"], False))
                for l in pkg.rel_by_type(m["part"], "/slideLayout"):
                    targets.append((l["part"], "layout %s" % l["part"], False))
        # Gather texts
        jobs = []  # (part, label, hidden, chunk_ts, chunk_text, para_full)
        texts = []
        for part, label, hidden in targets:
            root = pkg.xml(part)
            for p in root.iter(q("a", "p")):
                chunks, full = paragraph_chunks(p)
                texts.append(full)
                for ts in chunks:
                    jobs.append((part, label, hidden, ts))
                    texts.append("".join(t.text or "" for t in ts))
        # Charts and SmartArt: count only
        other = {}
        for n in pkg.names:
            if not pkg.exists(n) or not n.endswith(".xml"):
                continue
            kind = "charts" if n.startswith("ppt/charts/") else "smartart" if n.startswith("ppt/diagrams/") else None
            if kind and "/_rels/" not in n:
                root = pkg.xml(n)
                texts.append(" ".join((e.text or "") for e in root.iter(q("a", "t"), q("c", "v"))))
                other.setdefault(kind, []).append(len(texts) - 1)
        spans = find_all(pattern, flags, repl, a.regex, texts, a.timeout)
        # Map results back: paragraphs' full text gives total, chunks give replaceable
        idx, replaced, skipped, hits, touched = 0, 0, 0, [], set()
        for part, label, hidden in targets:
            root = pkg.xml(part)
            for p in root.iter(q("a", "p")):
                chunks, full = paragraph_chunks(p)
                total_here = len(spans[idx])
                idx += 1
                got = 0
                for ts in chunks:
                    ms = spans[idx]
                    idx += 1
                    if ms:
                        joined = "".join(t.text or "" for t in ts)
                        for st, en, _ in ms:
                            hits.append({"where": label, "hidden_slide": hidden,
                                         "text": pk.clean_text(joined[max(0, st - 30):en + 30], 120),
                                         "runs": sum(1 for t, (b0, b1) in zip(ts, _bounds(ts)) if b0 < en and b1 > st)})
                        if not a.dry_run:
                            apply(ts, ms)
                            touched.add(part)
                        got += len(ms)
                replaced += got
                skipped += max(0, total_here - got)
        not_replaced = {k: sum(len(spans[i]) for i in v) for k, v in other.items() if sum(len(spans[i]) for i in v)}
        res = {"ok": True, "input": a.input, "output": None if a.dry_run else a.output,
               "replaced": replaced, "skipped_across_breaks_or_fields": skipped,
               "not_replaced": not_replaced, "matches": hits[:200], "error": None,
               "hidden_slides_touched": sorted({h["where"] for h in hits if h["hidden_slide"]})}
        if a.expect is not None and replaced != a.expect:
            res.update(ok=False, output=None,
                       error="expected %d replacement(s), found %d; nothing written" % (a.expect, replaced))
            pk.emit(res, a.json, lambda r: print(r["error"]))
            return 1
        if replaced == 0:
            res.update(ok=False, output=None, error="no match; nothing written")
            pk.emit(res, a.json, lambda r: print(r["error"]))
            return 1
        if not a.dry_run:
            for part in touched:
                pkg.set_xml(part, pkg.xml(part))
            pkg.save(a.output, force=a.force)
        pk.emit(res, a.json, lambda r: print("%s %d replacement(s); skipped %d; not replaced %s" % (
            "would make" if a.dry_run else "made", r["replaced"], r["skipped_across_breaks_or_fields"],
            r["not_replaced"] or "none")))
        return 0
    except (pk.Unsupported, pk.Refused) as exc:
        return pk.fail(exc, a.json)


def _bounds(ts):
    out, pos = [], 0
    for t in ts:
        n = len(t.text or "")
        out.append((pos, pos + n))
        pos += n
    return out


if __name__ == "__main__":
    sys.exit(main())
