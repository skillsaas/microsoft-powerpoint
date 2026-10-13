#!/usr/bin/env python3
"""Make a share-ready copy of a deck, or check whether a deck is ready to share.

Usage:
  python pptx_sanitize.py IN --check-clean [--json]
  python pptx_sanitize.py IN OUT [--all | category flags] [--dry-run] [--force] [--json]

Category flags (each removes what inspect_pptx.py reports under the matching risk):
  --remove-notes          speaker notes slides (notesSlide parts); the notes master stays
  --remove-comments       legacy and modern comments, comment authors
  --remove-hidden-slides  slides with show="0" and slide parts left in the package but not in
                          the slide list (a common python-pptx delete leftover), plus their
                          section entries, stale section ids, custom show entries and links
                          from other slides that jump to them
  --remove-macros         vbaProject.bin and custom ribbon UI; .pptm/.ppsm/.potm become
                          .pptx/.ppsx/.potx (OUT must use the new extension)
  --remove-active         OLE objects, ActiveX controls, actions that run programs, macros or
                          OLE verbs, links with schemes other than http, https and mailto,
                          and shapes that pull external images, media or objects
  --remove-chart-workbooks  the workbook embedded behind each chart (the chart keeps its
                          cached values but "Edit Data" stops working); also external
                          chart data links
  --remove-hidden-shapes  shapes hidden in the selection pane and shapes placed off the slide
  --strip-metadata        author, last modified by, company, manager, hyperlink base, the
                          TitlesOfParts list in app.xml (every slide title, hidden ones
                          included), custom properties, thumbnail, tags, revision info
  --all                   every flag above

What it does not do: it is not redaction. Text on visible slides, text inside pictures,
cropped-away image areas, alt text, slide masters and layouts, embedded fonts and media files
stay. Review the rendered PDF before sending.

After writing, OUT is re-inspected. Exit 1 if any requested category still shows a risk
(with --all: if any risk remains at all), and the remaining items are listed.

Exit codes: 0 clean (or check passed), 1 residue left or --check-clean found risks,
2 unsupported input or output problem (exists, same as input, wrong extension),
3 safety refusal from the package reader.
"""
import argparse
import os
import sys
from collections import deque

import pptpkg as pk
from pptpkg import NS, q
import inspect_pptx as ins

CATS = ["notes", "comments", "hidden_slides", "macros", "active", "chart_workbooks",
        "hidden_shapes", "metadata"]
# risk names from inspect_pptx.analyze that each category must clear
CLEARS = {
    "notes": {"speaker-notes"},
    "comments": {"comments", "comment-authors"},
    "hidden_slides": {"hidden-slides", "unlisted-slides"},
    "macros": {"vba-macros", "custom-ribbon-ui", "macro-enabled-type"},
    "active": {"ole-objects", "activex-controls", "action", "external-link", "mismatched-link"},
    "chart_workbooks": {"embedded-files"},
    "hidden_shapes": {"hidden-shapes", "offslide-shapes"},
    "metadata": {"metadata", "tags"},
}
EXT_FOR = {"pptx": ".pptx", "ppsx": ".ppsx", "potx": ".potx", "pptm": ".pptm", "ppsm": ".ppsm", "potm": ".potm"}


def reachable(pkg):
    seen, todo = set(), deque([None])
    while todo:
        part = todo.popleft()
        for r in pkg.rels(part):
            if r["external"] or not r["part"] or r["part"] in seen:
                continue
            if pkg.exists(r["part"]):
                seen.add(r["part"])
                todo.append(r["part"])
    return seen


def xml_parts(pkg, prefixes=("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/",
                             "ppt/notesSlides/", "ppt/notesMasters/")):
    return [n for n in pkg.names if pkg.exists(n) and n.endswith(".xml") and n.startswith(prefixes)
            and "/_rels/" not in n]


def drop_unused_rel(pkg, part, rid):
    root = pkg.xml(part)
    if root is not None and not pk.r_attr_users(root, rid):
        pkg.drop_rel(part, rid)


def remove_shape(el, log, what):
    top = pk.top_shape(el)
    if top is None:
        # inside mc:AlternateContent that is itself a child of spTree?
        return False
    top.getparent().remove(top)
    log.append(what)
    return True


def do_notes(pkg, log):
    for s in pkg.slides():
        if s["part"]:
            for r in pkg.rel_by_type(s["part"], "/notesSlide"):
                pkg.drop_rel(s["part"], r["id"])
                log.append("notes of slide %d" % s["index"])


def do_comments(pkg, log):
    for s in pkg.slides():
        if s["part"]:
            for r in pkg.rel_by_type(s["part"], "/comments"):
                pkg.drop_rel(s["part"], r["id"])
                log.append("comments on slide %d" % s["index"])
    for r in pkg.rels(pkg.main):
        if r["type"].endswith(("/commentAuthors", "/authors")):
            pkg.drop_rel(pkg.main, r["id"])
            log.append("comment authors")


def remove_slides(pkg, gone, log, extra_parts=(), label="slide"):
    """Remove slides completely: sldId, presentation relationship, section and custom show
    entries (and stale section ids), links from any other part that jump to them, and the
    parts themselves (notes, charts and media go with the garbage collection that follows).

    gone: slide dicts from pkg.slides(); extra_parts: slide parts not in the slide list.
    """
    pres = pkg.xml(pkg.main)
    listed = pkg.slides()
    ids = {s["id"] for s in gone}
    rids = {s["rid"] for s in gone}
    parts = {s["part"] for s in gone} | set(extra_parts)
    known = {s["id"] for s in listed} - ids
    lst = pres.find("p:sldIdLst", NS)
    for sid in list(lst if lst is not None else []):
        if sid.get("id") in ids:
            lst.remove(sid)
    for el in list(pres.iter(q("p14", "sldId"))):
        if el.get("id") in ids or el.get("id") not in known:  # also stale ids
            el.getparent().remove(el)
    for el in list(pres.iter(q("p", "sld"))):  # custom show entries
        if el.get(q("r", "id")) in rids:
            el.getparent().remove(el)
    pkg.set_xml(pkg.main, pres)
    for r in pkg.rels(pkg.main):
        if r["part"] in parts:
            rids.add(r["id"])
    for rid in rids:
        pkg.drop_rel(pkg.main, rid)
    # Links in any remaining part (slides, notes, layouts, masters) that jump to a removed slide
    for part in ins.xml_part_names(pkg):
        if part in parts:
            continue
        hits = [r for r in pkg.rels(part) if r["type"].endswith("/slide") and r["part"] in parts]
        if not hits:
            continue
        root = pkg.xml(part)
        for r in hits:
            for el in pk.r_attr_users(root, r["id"]):
                el.getparent().remove(el)
            pkg.drop_rel(part, r["id"])
            log.append("link to removed slide in %s" % part)
        pkg.set_xml(part, root)
    log.extend("%s %s" % (label, s["index"]) for s in gone)
    log.extend("unlisted slide part %s" % n for n in sorted(extra_parts))
    for n in parts:  # drop the parts even if something unexpected still points at them
        pkg.remove(n)


def do_hidden_slides(pkg, log):
    listed = pkg.slides()
    gone = [s for s in listed if s["hidden"]]
    unlisted = set(ins.unlisted_slides(pkg, {s["part"] for s in listed}))
    if gone or unlisted:
        remove_slides(pkg, gone, log, unlisted, label="hidden slide")


def do_macros(pkg, log):
    for part, r in pkg.all_rels():  # by relationship type, whatever the part is called
        t = r["type"]
        if t.endswith(("/vbaProject", "/vbaProjectSignature", "/vbaProjectSignatureAgile",
                       "/vbaProjectSignatureV3")) or "ui/extensibility" in t or "customUI" in t:
            pkg.drop_rel(part, r["id"])
            log.append("macro or ribbon part %s" % (r["part"] or r["target"]))
    main_ct = pkg.content_type(pkg.main)
    if main_ct in pk.MACRO_FREE:
        for o in pkg.ct.findall("ct:Override", NS):
            if o.get("PartName", "").lstrip("/") == pkg.main:
                o.set("ContentType", pk.MACRO_FREE[main_ct])
        pkg.set_xml("[Content_Types].xml", pkg.ct)
        pkg.kind = pk.MAIN_TYPES[pk.MACRO_FREE[main_ct]]
        log.append("main part content type made macro-free (%s)" % pkg.kind)
    # vbaProject.bin also has a Default for .bin in some files; that is harmless once unreferenced


def do_active(pkg, log):
    for part in ins.xml_part_names(pkg):
        root = pkg.xml(part)
        rmap = {r["id"]: r for r in pkg.rels(part)}
        touched = []
        for tag in ("hlinkClick", "hlinkHover", "hlinkMouseOver"):
            for h in list(root.iter(q("a", tag))):
                action = h.get("action", "").split("?", 1)[0]
                rel = rmap.get(h.get(q("r", "id")))
                risky = action in ins.ACTION_RISK
                if rel is not None and rel["external"]:
                    if ins.lk.classify(rel["target"])[2] or ins.lk.mismatch(ins._link_text(h), rel["target"]):
                        risky = True  # unsafe scheme, or text that shows a different site
                if risky:
                    h.getparent().remove(h)
                    if rel is not None:
                        touched.append(rel["id"])
                    log.append("%s action %s in %s" % (tag, action or "link", part))
        for o in list(root.iter(q("p", "oleObj"))):
            if remove_shape(o, log, "OLE object in %s" % part):
                touched.extend(v for k, v in o.attrib.items() if k.startswith("{%s}" % pk.R_NS))
        for c in list(root.iter(q("p", "controls"))):
            c.getparent().remove(c)
            log.append("ActiveX controls in %s" % part)
        for rid, r in rmap.items():
            if r["external"] and not r["type"].endswith("/hyperlink"):
                for el in pk.r_attr_users(root, rid):
                    if not remove_shape(el, log, "shape using external %s in %s" % (r["type"].rsplit("/", 1)[-1], part)):
                        for k in [k for k, v in el.attrib.items() if v == rid]:
                            del el.attrib[k]
                touched.append(rid)
        pkg.set_xml(part, root)
        for rid in set(touched):
            drop_unused_rel(pkg, part, rid)
        for r in pkg.rels(part):  # ActiveX / OLE rels left without users
            if r["type"].endswith(("/control", "/oleObject", "/package")) and not pk.r_attr_users(root, r["id"]):
                pkg.drop_rel(part, r["id"])
    # External targets left in any .rels file (chart data, theme pictures, SmartArt, presProps,
    # package level): remove what uses them, or the reference, then the relationship.
    for src, r in pkg.all_rels():
        if not r["external"] or r["type"].endswith("/hyperlink"):
            continue
        root = pkg.xml(src) if src and src.endswith(".xml") and pkg.exists(src) else None
        if root is not None:
            for el in pk.r_attr_users(root, r["id"]):
                if src.startswith("ppt/charts/") and el.tag == q("c", "externalData"):
                    el.getparent().remove(el)
                elif not remove_shape(el, log, "shape using external %s in %s" % (r["type"].rsplit("/", 1)[-1], src)):
                    for k in [k for k, v in el.attrib.items() if v == r["id"]]:
                        del el.attrib[k]
            pkg.set_xml(src, root)
        pkg.drop_rel(src, r["id"])
        log.append("external %s link in %s" % (r["type"].rsplit("/", 1)[-1], src or "package"))


def do_chart_workbooks(pkg, log):
    for n in [n for n in pkg.names if pkg.exists(n) and n.startswith("ppt/charts/") and n.endswith(".xml")
              and "/_rels/" not in n]:
        root = pkg.xml(n)
        for ed in list(root.iter(q("c", "externalData"))):
            rid = ed.get(q("r", "id"))
            ed.getparent().remove(ed)
            pkg.set_xml(n, root)
            pkg.drop_rel(n, rid)
            log.append("chart workbook behind %s" % n)


def do_hidden_shapes(pkg, log, cx, cy):
    for s in pkg.slides():
        if not s["part"]:
            continue
        root = pkg.xml(s["part"])
        for cnv in list(root.iter(q("p", "cNvPr"))):
            if cnv.get("hidden") in ("1", "true"):
                remove_shape(cnv, log, "hidden shape '%s' on slide %d" % (pk.clean_text(cnv.get("name", ""), 80), s["index"]))
        for shp in pk.offslide_shapes(root, cx, cy):
            shp.getparent().remove(shp)
            log.append("off-slide shape on slide %d" % s["index"])
        pkg.set_xml(s["part"], root)


def do_metadata(pkg, log):
    if pkg.exists("docProps/core.xml"):
        core = pkg.xml("docProps/core.xml")
        for el in list(core):
            if isinstance(el.tag, str) and el.tag.split("}")[1] in pk.CORE_STRIP:
                if (el.text or "").strip():
                    log.append("core property %s" % el.tag.split("}")[1])
                core.remove(el)
        pkg.set_xml("docProps/core.xml", core)
    if pkg.exists("docProps/app.xml"):
        app = pkg.xml("docProps/app.xml")
        for el in list(app):
            if isinstance(el.tag, str) and el.tag.split("}")[1] in pk.APP_STRIP:
                if (el.text or "").strip() or len(el):
                    log.append("app property %s" % el.tag.split("}")[1])
                app.remove(el)
        pkg.set_xml("docProps/app.xml", app)
    for r in pkg.rels(None):
        if r["type"].endswith(("/custom-properties", "/thumbnail")):
            pkg.drop_rel(None, r["id"])
            log.append("package part %s" % r["part"])
    for r in pkg.rels(pkg.main):
        if r["type"].endswith("/revisionInfo"):
            pkg.drop_rel(pkg.main, r["id"])
            log.append("revision info")
    for part in xml_parts(pkg) + [pkg.main]:
        root = pkg.xml(part)
        hit = False
        for cd in list(root.iter(q("p", "custDataLst"))):
            cd.getparent().remove(cd)
            hit = True
        if hit:
            pkg.set_xml(part, root)
            for r in pkg.rels(part):
                if r["type"].endswith("/tags"):
                    pkg.drop_rel(part, r["id"])
            log.append("tags in %s" % part)


def collect_garbage(pkg, before, log):
    after = reachable(pkg)
    for n in sorted(before - after):
        pkg.remove(n)
        log.append("unreferenced part %s" % n)


def check(pkg_path, cats):
    res = ins.analyze(pk.Package(pkg_path), want_text=False)
    names = set().union(*(CLEARS[c] for c in cats)) if cats else None
    if names is not None and set(cats) == set(CATS):
        names = None  # --all: every risk counts
    left = [r for r in res["risks"] if names is None or r["risk"] in names]
    return left, res


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("input")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--check-clean", action="store_true")
    ap.add_argument("--all", action="store_true")
    for c in CATS:
        flag = "--strip-metadata" if c == "metadata" else "--remove-" + c.replace("_", "-")
        ap.add_argument(flag, dest=c, action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true", help="replace OUT if it exists (no backup)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    cats = CATS[:] if a.all else [c for c in CATS if getattr(a, c)]
    try:
        if a.check_clean:
            left, res = check(a.input, CATS)
            out = {"ok": not left, "file": a.input, "kind": res["kind"], "residue": left,
                   "clean_scope": res["clean_scope"],
                   "error": None if not left else "deck is not ready to share",
                   "note": "ok means these checks passed, not that the deck is safe to share: "
                           "review clean_scope.not_checked with the user."}
            pk.emit(out, a.json, lambda r: print("clean" if r["ok"] else
                                                 "NOT clean:\n" + "\n".join("  - %s" % x for x in r["residue"])))
            return 0 if not left else 1
        if not a.output:
            raise pk.Unsupported("give OUT, or use --check-clean")
        if not cats:
            raise pk.Unsupported("choose --all or at least one category flag")
        pkg = pk.Package(a.input)
        res0 = ins.analyze(pkg, want_text=False)
        cx = res0.get("slide_size_emu", {}).get("cx", 0)
        cy = res0.get("slide_size_emu", {}).get("cy", 0)
        before = reachable(pkg)
        log = []
        for c in cats:
            if c == "hidden_shapes":
                do_hidden_shapes(pkg, log, cx, cy)
            else:
                globals()["do_" + c](pkg, log)
        collect_garbage(pkg, before, log)
        want_ext = EXT_FOR[pkg.kind]
        if os.path.splitext(a.output)[1].lower() != want_ext:
            raise pk.Unsupported("OUT must end in %s for a %s package (PowerPoint refuses a mismatch)"
                                 % (want_ext, pkg.kind))
        if a.dry_run:
            out = {"ok": True, "dry_run": True, "would_remove": log, "error": None}
            pk.emit(out, a.json, lambda r: print("\n".join(r["would_remove"]) or "nothing to remove"))
            return 0
        pkg.save(a.output, force=a.force)
        left, res = check(a.output, cats)
        out = {"ok": not left, "input": a.input, "output": a.output, "kind": res["kind"],
               "removed": log, "remaining": left, "clean_scope": res["clean_scope"],
               "error": None if not left else "risks remain after sanitizing",
               "note": "Not redaction: visible slide text, pictures, masters and media were kept."}
        pk.emit(out, a.json, lambda r: print("removed %d item(s); %s" % (
            len(r["removed"]), "clean" if r["ok"] else "REMAINING: %s" % r["remaining"])))
        return 0 if not left else 1
    except (pk.Unsupported, pk.Refused) as exc:
        return pk.fail(exc, a.json)


if __name__ == "__main__":
    sys.exit(main())
