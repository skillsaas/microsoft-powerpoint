#!/usr/bin/env python3
"""Delete, hide, unhide or move slides without python-pptx private attributes.

Usage:
  python pptx_slides.py IN OUT [--delete 3,5-6] [--hide 2] [--unhide 4] [--move FROM:TO ...]
                        [--expect-slides N] [--dry-run] [--force] [--json]

Numbers are 1-based positions in show order (p:sldIdLst) as inspect_pptx.py prints them.
--delete, --hide and --unhide use the input deck's numbers. --move FROM:TO uses the order after
the deletes and after earlier moves, left to right. To avoid counting, delete in one run and
move in a second run on its output.

Delete is complete: the sldId, the presentation relationship, section and custom show entries,
links from other slides, notes or masters that jump to it, the slide part, and parts only it
used (notes, charts, media, comments) are removed, and the package is checked afterwards.
Move keeps sections contiguous: a moved slide joins the section of the slide it lands after
(or the first section when it lands first).

Exit codes: 0 written (or dry run), 1 --expect-slides does not match the result or the result
fails the consistency check (nothing written), 2 bad arguments, unsupported input or output
problem, 3 safety refusal from the package reader.
"""
import argparse
import sys

import pptpkg as pk
from pptpkg import NS, q
import pptx_sanitize as ps


def numbers(spec, n, flag):
    out = []
    for bit in (spec or "").split(","):
        bit = bit.strip()
        if not bit:
            continue
        try:
            if "-" in bit:
                a, b = (int(x) for x in bit.split("-", 1))
                out.extend(range(a, b + 1))
            else:
                out.append(int(bit))
        except ValueError:
            raise pk.Unsupported("%s: not a slide number list: %s" % (flag, bit))
    bad = [i for i in out if i < 1 or i > n]
    if bad:
        raise pk.Unsupported("%s: slide numbers out of range 1..%d: %s" % (flag, n, bad))
    return out


def consistency(pkg):
    """Problems PowerPoint would repair: dangling rels, missing overrides, stale section ids."""
    probs = []
    for src, r in pkg.all_rels():
        if not r["external"] and r["part"] and not pkg.exists(r["part"]):
            probs.append("dangling relationship %s -> %s" % (src or "package", r["part"]))
    for o in pkg.ct.findall("ct:Override", NS):
        if not pkg.exists(o.get("PartName", "").lstrip("/")):
            probs.append("content type for missing part %s" % o.get("PartName"))
    ids = {s["id"] for s in pkg.slides()}
    for sec in pkg.sections():
        probs.extend("section %s lists missing slide id %s" % (sec["name"], i) for i in sec["slide_ids"] if i not in ids)
    return probs


def move(pkg, frm, to, log):
    pres = pkg.xml(pkg.main)
    lst = pres.find("p:sldIdLst", NS)
    items = list(lst)
    if not (1 <= frm <= len(items) and 1 <= to <= len(items)):
        raise pk.Unsupported("--move %d:%d out of range 1..%d (after deletes)" % (frm, to, len(items)))
    el = items[frm - 1]
    lst.remove(el)
    lst.insert(to - 1, el)
    order = [e.get("id") for e in lst]
    sid = el.get("id")
    secs = [sec for ext in pres.iter(q("p", "ext")) if ext.get("uri") == pk.SECTION_EXT
            for sec in ext.iter(q("p14", "section"))]
    if secs:
        for sec in secs:
            for x in list(sec.iter(q("p14", "sldId"))):
                if x.get("id") == sid:
                    x.getparent().remove(x)
        pos = order.index(sid)
        target = secs[0]
        if pos > 0:
            before = order[pos - 1]
            target = next((sec for sec in secs if any(x.get("id") == before for x in sec.iter(q("p14", "sldId")))), secs[0])
        lst_el = target.find("p14:sldIdLst", NS)
        if lst_el is None:
            lst_el = target.makeelement(q("p14", "sldIdLst"), {})
            target.append(lst_el)
        new = lst_el.makeelement(q("p14", "sldId"), {"id": sid})
        lst_el.append(new)
        for sec in secs:  # keep each section in show order
            l2 = sec.find("p14:sldIdLst", NS)
            if l2 is not None:
                kids = sorted(list(l2), key=lambda x: order.index(x.get("id")) if x.get("id") in order else 1e9)
                for k in kids:
                    l2.remove(k)
                    l2.append(k)
    pkg.set_xml(pkg.main, pres)
    log.append("moved slide %d to position %d" % (frm, to))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--delete")
    ap.add_argument("--hide")
    ap.add_argument("--unhide")
    ap.add_argument("--move", action="append", default=[], metavar="FROM:TO")
    ap.add_argument("--expect-slides", type=int)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true", help="replace OUT if it exists (no backup)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    try:
        if not (a.delete or a.hide or a.unhide or a.move):
            raise pk.Unsupported("nothing to do: give --delete, --hide, --unhide or --move")
        pkg = pk.Package(a.input)
        slides = pkg.slides()
        n = len(slides)
        dele = numbers(a.delete, n, "--delete")
        hide = numbers(a.hide, n, "--hide")
        unhide = numbers(a.unhide, n, "--unhide")
        if set(hide) & set(unhide):
            raise pk.Unsupported("a slide cannot be in both --hide and --unhide")
        if len(set(dele)) == n:
            raise pk.Unsupported("refusing to delete every slide")
        moves = []
        for m in a.move:
            try:
                f, t = (int(x) for x in m.split(":", 1))
            except ValueError:
                raise pk.Unsupported("--move takes FROM:TO, for example 5:1")
            moves.append((f, t))
        log = []
        before = ps.reachable(pkg)
        for i in hide + unhide:
            s = slides[i - 1]
            root = pkg.xml(s["part"])
            if i in hide:
                root.set("show", "0")
                log.append("hid slide %d" % i)
            elif "show" in root.attrib:
                del root.attrib["show"]
                log.append("unhid slide %d" % i)
            pkg.set_xml(s["part"], root)
        if dele:
            ps.remove_slides(pkg, [slides[i - 1] for i in sorted(set(dele))], log, label="deleted slide")
            ps.collect_garbage(pkg, before, log)
        for f, t in moves:
            move(pkg, f, t, log)
        result = len(pkg.slides())
        probs = consistency(pkg)
        res = {"ok": True, "input": a.input, "output": None if a.dry_run else a.output,
               "slides_before": n, "slides_after": result, "changes": log, "problems": probs, "error": None}
        if probs:
            res.update(ok=False, output=None, error="result would need repair; nothing written")
        elif a.expect_slides is not None and result != a.expect_slides:
            res.update(ok=False, output=None,
                       error="expected %d slides after the change, got %d; nothing written" % (a.expect_slides, result))
        if not res["ok"]:
            pk.emit(res, a.json, lambda r: print(r["error"]))
            return 1
        if not a.dry_run:
            pkg.save(a.output, force=a.force)
        pk.emit(res, a.json, lambda r: print("%s: %d -> %d slides; %s" % (
            "dry run" if a.dry_run else "wrote " + a.output, n, result, "; ".join(log))))
        return 0
    except (pk.Unsupported, pk.Refused) as exc:
        return pk.fail(exc, a.json)


if __name__ == "__main__":
    sys.exit(main())
