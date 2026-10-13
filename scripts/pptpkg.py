"""Package layer shared by the microsoft-powerpoint scripts (not a public API).

A .pptx is an OPC zip: [Content_Types].xml, _rels/.rels, ppt/presentation.xml and one part per
slide, layout, master, notes slide, chart, media file and embedding, tied together by .rels files.
This module opens such a zip defensively and writes a modified copy without touching parts that
were not changed.

Defences (each one has a test in tests/smoke_test.py):
  - Zip caps before any XML is parsed: MAX_MEMBERS entries, MAX_TOTAL uncompressed bytes,
    MAX_PART bytes for one XML part, and a compression-ratio cap for XML parts.
  - Duplicate member names and names equal after case folding are refused, because two readers
    could pick different copies and the part we check would not be the part PowerPoint shows.
  - Member names with "..", a leading "/" or a backslash are refused.
  - XML is parsed with lxml with entity resolution, DTD loading and network access off; a
    DOCTYPE is refused outright (OOXML never needs one).
  - Encrypted decks (an OLE2 compound file holding EncryptionInfo) and legacy binary .ppt are
    detected by their magic bytes and reported as unsupported, not as corrupt.
  - Writes never replace an existing path: the copy is written to a temp file in the target
    folder and hard-linked into place, which fails if the target appeared meanwhile.
"""
import os
import posixpath
import re
import sys
import tempfile
import zipfile

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    sys.stderr.write("lxml is required: pip install -r scripts/requirements.txt\n")
    sys.exit(2)

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
}
R_NS = NS["r"]
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
RT_MS = "http://schemas.microsoft.com/office/"

# Main part content type -> short kind. Macro kinds can carry VBA.
MAIN_TYPES = {
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml": "pptx",
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml": "ppsx",
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml": "potx",
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml": "pptm",
    "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml": "ppsm",
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml": "potm",
}
MACRO_FREE = {
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml":
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml":
        "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml":
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
}
# Section list extension URI used by PowerPoint 2010 and later.
SECTION_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"

# Document properties the share-check reports and --strip-metadata removes (one list for both).
CORE_STRIP = ("creator", "lastModifiedBy", "lastPrinted", "keywords", "description", "subject", "category",
              "contentStatus", "identifier", "title", "revision", "language", "version")
APP_STRIP = ("Company", "Manager", "HyperlinkBase", "Template", "TotalTime", "HeadingPairs", "TitlesOfParts")

MAX_MEMBERS = 10000
MAX_TOTAL = 1024 * 1024 * 1024  # 1 GiB uncompressed (decks carry video)
MAX_PART = 64 * 1024 * 1024     # one XML part
MAX_XML_RATIO = 200             # uncompressed / compressed for XML parts

# C0 and C1 controls, zero-width characters, bidi embeddings, overrides and isolates, BOM.
_UNSAFE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\u061c\u200b-\u200f"
                     "\u202a-\u202e\u2060-\u2064\u2066-\u2069\ufeff]")

EXIT_OK, EXIT_CHECK, EXIT_UNSUPPORTED, EXIT_REFUSED = 0, 1, 2, 3


class Unsupported(Exception):
    """Input is not something we can process (exit 2)."""


class Refused(Exception):
    """Safety refusal (exit 3)."""


def q(prefix, local):
    return "{%s}%s" % (NS[prefix], local)


def clean_text(s, limit=2000):
    """Make file text safe to print: drop controls and invisible characters, cap the length."""
    s = _UNSAFE.sub("", s or "").replace("\r", " ")
    if len(s) > limit:
        s = s[:limit] + "...[truncated %d chars]" % (len(s) - limit)
    return s


def sniff(path):
    with open(path, "rb") as fh:
        head = fh.read(8)
    if head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06"):
        return "zip"
    if head == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "ole2"
    return "other"


def ole2_kind(path):
    """Tell an encrypted OOXML file from a legacy .ppt by looking for stream names (UTF-16LE)."""
    with open(path, "rb") as fh:
        blob = fh.read(4 * 1024 * 1024)
    if "EncryptionInfo".encode("utf-16-le") in blob:
        return "encrypted"
    if "PowerPoint Document".encode("utf-16-le") in blob:
        return "ppt"
    return "ole2"


_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False,
                          huge_tree=False, remove_blank_text=False)


def parse_xml(data, name="part"):
    if b"<!DOCTYPE" in data[:4096] or b"<!ENTITY" in data:
        raise Refused("%s contains a DOCTYPE or ENTITY declaration" % name)
    try:
        return etree.fromstring(data, _PARSER)
    except etree.XMLSyntaxError as exc:
        raise Unsupported("%s is not well-formed XML: %s" % (name, exc))


def rels_name(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def resolve(source_part, target):
    """Resolve a relationship target relative to the part that owns the .rels file."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def check_zip(zf):
    """Zip caps and name checks shared by every script (also used for .odp before rendering).

    Caps are read from the module at call time, so tests can lower them.
    """
    infos = zf.infolist()
    if len(infos) > MAX_MEMBERS:
        raise Refused("too many zip entries (%d > %d)" % (len(infos), MAX_MEMBERS))
    seen, seen_low, total = set(), set(), 0
    for info in infos:
        n = info.filename
        if n in seen:
            raise Refused("duplicate zip entry: %s" % clean_text(n, 200))
        if n.lower() in seen_low:
            raise Refused("zip entries differ only by case: %s" % clean_text(n, 200))
        seen.add(n)
        seen_low.add(n.lower())
        if ".." in n.split("/") or n.startswith("/") or "\\" in n:
            raise Refused("unsafe zip entry name: %s" % clean_text(n, 200))
        total += info.file_size
        if n.endswith((".xml", ".rels")):
            if info.file_size > MAX_PART:
                raise Refused("XML part too large: %s" % n)
            if info.compress_size and info.file_size / max(info.compress_size, 1) > MAX_XML_RATIO \
                    and info.file_size > 1024 * 1024:
                raise Refused("suspicious compression ratio in %s" % n)
    if total > MAX_TOTAL:
        raise Refused("uncompressed size %d exceeds cap %d" % (total, MAX_TOTAL))
    return infos


class Package:
    """Read-only view of a presentation package, with edits kept in memory until save()."""

    def __init__(self, path):
        self.path = path
        kind = sniff(path)
        if kind == "ole2":
            k = ole2_kind(path)
            if k == "encrypted":
                raise Unsupported("the file is password-protected (encrypted OOXML)")
            if k == "ppt":
                raise Unsupported("legacy binary .ppt: convert a copy to .pptx first (pptx_render.py --to pptx)")
            raise Unsupported("OLE2 compound file that is not a presentation")
        if kind != "zip":
            raise Unsupported("not a zip package (.pptx, .pptm, .ppsx, .potx)")
        try:
            self.zf = zipfile.ZipFile(path)
        except zipfile.BadZipFile as exc:
            raise Unsupported("corrupt zip: %s" % exc)
        infos = check_zip(self.zf)
        self.names = [i.filename for i in infos]
        self._lower = {n.lower(): n for n in self.names}
        self.changed = {}    # name -> bytes (or etree root)
        self.removed = set()
        self._xml = {}
        if "[Content_Types].xml" not in self.names:
            raise Unsupported("no [Content_Types].xml: not an OPC package")
        self.ct = self.xml("[Content_Types].xml")
        self.main = self._find_main()
        self.kind = self.content_type(self.main) and MAIN_TYPES.get(self.content_type(self.main))
        if not self.kind:
            raise Unsupported("main part is not a PowerPoint presentation (content type %s)"
                              % self.content_type(self.main))

    # ----- reading
    def exists(self, name):
        return name in self.names and name not in self.removed

    def read(self, name):
        if name in self.changed:
            v = self.changed[name]
            return v if isinstance(v, bytes) else etree.tostring(v, xml_declaration=True,
                                                                  encoding="UTF-8", standalone=True)
        return self.zf.read(name)

    def xml(self, name):
        if name not in self._xml:
            if not self.exists(name):
                return None
            self._xml[name] = parse_xml(self.zf.read(name), name)
        return self._xml[name]

    def content_type(self, name):
        for o in self.ct.findall("ct:Override", NS):
            if o.get("PartName", "").lstrip("/") == name:
                return o.get("ContentType")
        ext = name.rsplit(".", 1)[-1].lower()
        for d in self.ct.findall("ct:Default", NS):
            if d.get("Extension", "").lower() == ext:
                return d.get("ContentType")
        return None

    def rels(self, part):
        """List of dicts: id, type, target (raw), external (bool), part (resolved name or None)."""
        root = self.xml(rels_name(part)) if part else self.xml("_rels/.rels")
        out = []
        if root is None:
            return out
        for r in root.findall("pr:Relationship", NS):
            ext = r.get("TargetMode") == "External"
            tgt = r.get("Target", "")
            out.append({"id": r.get("Id"), "type": r.get("Type", ""), "target": tgt,
                        "external": ext, "part": None if ext else resolve(part or "", tgt),
                        "el": r})
        return out

    def all_rels(self):
        """(source part or None, rel) for every .rels file in the package."""
        out = []
        for n in self.names:
            if not n.endswith(".rels") or not self.exists(n):
                continue
            if n == "_rels/.rels":
                src = None
            else:
                d, b = posixpath.split(n)
                src = posixpath.join(posixpath.dirname(d), b[:-5])
            for r in self.rels(src):
                out.append((src, r))
        return out

    def rel_by_type(self, part, suffix):
        return [r for r in self.rels(part) if r["type"].endswith(suffix)]

    def _find_main(self):
        for r in self.rels(None):
            if r["type"].endswith("/officeDocument"):
                return r["part"]
        raise Unsupported("no officeDocument relationship in _rels/.rels")

    def slides(self):
        """Slides in presentation order (p:sldIdLst), with id, rId, part, hidden flag."""
        pres = self.xml(self.main)
        rmap = {r["id"]: r for r in self.rels(self.main)}
        out = []
        lst = pres.find("p:sldIdLst", NS)
        for n, sid in enumerate(lst if lst is not None else [], 1):
            rid = sid.get(q("r", "id"))
            rel = rmap.get(rid)
            part = rel["part"] if rel and not rel["external"] else None
            root = self.xml(part) if part and self.exists(part) else None
            out.append({"index": n, "id": sid.get("id"), "rid": rid, "part": part,
                        "missing": root is None,
                        "hidden": root is not None and root.get("show") in ("0", "false")})
        return out

    def sections(self):
        pres = self.xml(self.main)
        out = []
        for ext in pres.iter(q("p", "ext")):
            if ext.get("uri") != SECTION_EXT:
                continue
            for sec in ext.iter(q("p14", "section")):
                ids = [s.get("id") for s in sec.iter(q("p14", "sldId"))]
                out.append({"name": clean_text(sec.get("name", ""), 200), "slide_ids": ids})
        return out

    # ----- writing
    def set_xml(self, name, root):
        self._xml[name] = root
        self.changed[name] = root

    def set_bytes(self, name, data):
        self._xml.pop(name, None)
        self.changed[name] = data

    def remove(self, name):
        if name in self.names:
            self.removed.add(name)
            self.changed.pop(name, None)
            self._xml.pop(name, None)
            for o in self.ct.findall("ct:Override", NS):
                if o.get("PartName", "").lstrip("/") == name:
                    self.ct.remove(o)
                    self.set_xml("[Content_Types].xml", self.ct)
            rn = rels_name(name)
            if rn in self.names and rn not in self.removed:
                self.removed.add(rn)
                self.changed.pop(rn, None)
                self._xml.pop(rn, None)

    def drop_rel(self, part, rid):
        name = rels_name(part) if part else "_rels/.rels"
        root = self.xml(name)
        if root is None:
            return
        for r in root.findall("pr:Relationship", NS):
            if r.get("Id") == rid:
                root.remove(r)
        self.set_xml(name, root)

    def save(self, out_path, force=False):
        out_path = os.path.abspath(out_path)
        if os.path.exists(out_path) and not force:
            raise Unsupported("output exists: %s (use --force to replace it, no backup is kept)" % out_path)
        if os.path.abspath(self.path) == out_path:
            raise Unsupported("output is the input file; write to a new path")
        d = os.path.dirname(out_path) or "."
        fd, tmp = tempfile.mkstemp(prefix=".pptx-", suffix=".tmp", dir=d)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
                # [Content_Types].xml first, as PowerPoint writes it.
                order = ["[Content_Types].xml"] + [n for n in self.names if n != "[Content_Types].xml"]
                for n in order:
                    if n in self.removed:
                        continue
                    info = self.zf.getinfo(n)
                    if n in self.changed:
                        zi = zipfile.ZipInfo(n, date_time=info.date_time)
                        zi.compress_type = zipfile.ZIP_DEFLATED
                        out.writestr(zi, self.read(n))
                    else:
                        out.writestr(info, self.zf.read(n))
                for n, v in self.changed.items():
                    if n not in self.names:
                        out.writestr(n, self.read(n))
            if force and os.path.exists(out_path):
                os.replace(tmp, out_path)
            else:
                os.link(tmp, out_path)  # fails if out_path appeared meanwhile
                os.unlink(tmp)
        except FileExistsError:
            os.unlink(tmp)
            raise Unsupported("output appeared while writing: %s" % out_path)
        except Exception:
            if os.path.exists(tmp):
                os.unlink(tmp)
            raise


def text_of(el):
    """Text of a DrawingML element: a:t joined per paragraph, a:br as newline, tabs kept."""
    paras = []
    for p in el.iter(q("a", "p")):
        bits = []
        for node in p.iter(q("a", "t"), q("a", "br")):
            bits.append("\n" if node.tag == q("a", "br") else (node.text or ""))
        paras.append("".join(bits))
    return "\n".join(paras)


def top_shape(el):
    """Climb to the element that is a direct child of p:spTree or p:grpSp (the whole shape)."""
    containers = (q("p", "spTree"), q("p", "grpSp"))
    cur = el
    while cur is not None and cur.getparent() is not None:
        if cur.getparent().tag in containers:
            return cur
        cur = cur.getparent()
    return None


def offslide_shapes(root, cx, cy):
    """Top-level shapes (direct children of p:spTree) placed entirely outside the slide.

    Only top-level shapes are checked: children of a group use the group's own coordinate
    space (a:chOff/a:chExt), so their offsets are not slide positions.
    """
    out = []
    tree = root.find("p:cSld/p:spTree", NS)
    if tree is None or not cx or not cy:
        return out
    for shp in tree:
        if not isinstance(shp.tag, str) or shp.tag in (q("p", "nvGrpSpPr"), q("p", "grpSpPr")):
            continue
        off = shp.find("./*/a:xfrm/a:off", NS)
        ext = shp.find("./*/a:xfrm/a:ext", NS)
        if off is None or ext is None:
            continue
        try:
            x, y = int(off.get("x", "0")), int(off.get("y", "0"))
            w, h = int(ext.get("cx", "0")), int(ext.get("cy", "0"))
        except ValueError:
            continue
        if x >= cx or y >= cy or (w > 0 and x + w <= 0) or (h > 0 and y + h <= 0):
            out.append(shp)
    return out


def r_attr_users(root, rid):
    """Elements in a part that reference relationship rid through any r:* attribute."""
    hits = []
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        for k, v in el.attrib.items():
            if k.startswith("{%s}" % R_NS) and v == rid:
                hits.append(el)
    return hits


def emit(result, as_json, human):
    import json
    if as_json:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        human(result)


def fail(exc, as_json):
    import json
    code = EXIT_REFUSED if isinstance(exc, Refused) else EXIT_UNSUPPORTED
    msg = clean_text(str(exc), 500)
    if as_json:
        print(json.dumps({"ok": False, "error": msg, "exit": code}))
    else:
        sys.stderr.write("error: %s\n" % msg)
    return code
