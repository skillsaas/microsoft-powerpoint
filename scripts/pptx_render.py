#!/usr/bin/env python3
"""Render a deck to PDF or per-slide PNG, or convert legacy .ppt/.odp to .pptx, with
LibreOffice Impress inside a bubblewrap sandbox.

Usage:
  python pptx_render.py --check [--json]
  python pptx_render.py IN OUT.pdf  [--notes-pages] [--timeout S] [--force] [--json]
  python pptx_render.py IN OUTDIR --to png [--dpi N] [--timeout S] [--json]
  python pptx_render.py IN OUT.pptx --to pptx [--timeout S] [--force] [--json]

Why a sandbox: a deck can link pictures, media and objects by URL or by local path. Tested on
LibreOffice 25.2.3.2: a plain `soffice --convert-to pdf` of a .pptx fetched an http:// picture
from a local server and embedded a file:// picture from /tmp into the PDF. Inside this
script's sandbox neither happens (tests/smoke_test.py, render canary tests).

Sandbox (bwrap): new network, PID, IPC, UTS and user namespaces (no network at all); a
filesystem with read-only /usr, fonts and the dynamic linker cache, an empty /tmp, and one
private work folder holding a copy of the input and a throwaway LibreOffice profile; no home
folder and no other user files; minimal environment; macros disabled in the profile
(MacroSecurityLevel 3, DisableMacrosExecution,
BlockUntrustedRefererLinks). The process group is killed at --timeout
(default 180 s). Files readable under /usr (public system files) are still visible.
Without bwrap the script refuses (exit 2): there is no unsandboxed mode and no flag that
skips the check. On macOS or Windows use PowerPoint's own export instead (see
references/fallbacks.md).

Limits: LibreOffice is not PowerPoint. Fonts, SmartArt, some charts, 3D models, morph and
other effects can render differently; the PDF shows LibreOffice's layout. Hidden slides:
LibreOffice 25.2.3.2 left them out of the PDF in our test. PNG export renders the PDF with
pdftoppm (poppler) in the same sandbox, one file per slide (slide-01.png ...), because
`soffice --convert-to png` writes only the first slide.

Input checks before LibreOffice sees the file, by content, not extension: 512 MB on disk;
.pptx family packages pass the package checks of pptpkg.py (entry count, uncompressed size,
duplicate names, DOCTYPE); .odp must be a zip with content.xml and passes the same zip caps;
.ppt/.pps/.pot must be OLE2 and not encrypted. A zip renamed .ppt, or OLE2 renamed .odp, is
refused. LibreOffice runs with an address-space cap (--memory-mb, default 4096) as well as
the timeout. --to pptx output is LibreOffice-written and keeps external links: inspect it.

Exit codes: 0 written, 1 LibreOffice or pdftoppm ran but produced no output, 2 setup problem
or unsupported input (no soffice, no bwrap, output exists, wrong extension), 3 safety refusal
(caps) or timeout.
"""
import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile

import pptpkg as pk

MAX_INPUT = 512 * 1024 * 1024
PKG_EXT = {".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm"}
LEGACY_EXT = {".ppt", ".pps", ".pot", ".odp"}

XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="UseOpenCL" oor:op="fuse"><value>false</value></prop></item>
</oor:items>
"""


def check_input(src, ext):
    """Caps and type checks by content, not extension, before LibreOffice sees the file."""
    kind = pk.sniff(src)
    if ext in PKG_EXT:
        pk.Package(src)  # zip caps, duplicate names, DOCTYPE, encrypted, legacy
        return
    if kind == "ole2":
        k = pk.ole2_kind(src)
        if k == "encrypted":
            raise pk.Unsupported("the file is password-protected")
        if ext == ".odp":
            raise pk.Unsupported("content is an OLE2 file, not OpenDocument; extension does not match")
        return  # legacy .ppt/.pps/.pot: binary format, size cap above, sandbox does the rest
    if kind == "zip":
        import zipfile
        with zipfile.ZipFile(src) as zf:
            pk.check_zip(zf)  # same entry, size and name caps as the .pptx path
            names = set(zf.namelist())
        if ext != ".odp":
            raise pk.Unsupported("content is a zip package, not a legacy binary %s; rename it .pptx or .odp" % ext)
        if "content.xml" not in names:
            raise pk.Unsupported("zip without content.xml is not an OpenDocument presentation")
        return
    raise pk.Unsupported("content does not match the %s extension" % ext)


def limit_memory(mb):
    """Address-space cap for LibreOffice (inherited through bwrap). Linux only."""
    def apply():
        try:
            import resource
            resource.setrlimit(resource.RLIMIT_AS, (mb * 1024 * 1024, mb * 1024 * 1024))
        except (ImportError, ValueError, OSError):
            pass
    return apply


def find_soffice():
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return os.path.realpath(p)
    for p in ("/usr/lib/libreoffice/program/soffice", "/opt/libreoffice/program/soffice"):
        if os.path.exists(p):
            return p
    return None


def bwrap_cmd(work):
    args = ["bwrap", "--unshare-all", "--die-with-parent", "--new-session",
            "--ro-bind", "/usr", "/usr", "--proc", "/proc", "--dev", "/dev", "--tmpfs", "/tmp", "--tmpfs", "/run",
            "--bind", work, "/work", "--chdir", "/work", "--clearenv",
            "--setenv", "HOME", "/work/home", "--setenv", "TMPDIR", "/work/tmp",
            "--setenv", "PATH", "/usr/bin:/bin", "--setenv", "LANG", "C.UTF-8",
            "--setenv", "SAL_USE_VCLPLUGIN", "svp"]
    for link in ("lib", "lib64", "bin", "sbin", "lib32"):
        p = "/" + link
        if os.path.islink(p):
            args += ["--symlink", os.readlink(p), p]
        elif os.path.isdir(p):
            args += ["--ro-bind", p, p]
    for e in ("/etc/fonts", "/etc/ld.so.cache", "/etc/ld.so.conf", "/etc/ld.so.conf.d", "/etc/alternatives",
              "/etc/libreoffice", "/etc/localtime"):
        if os.path.exists(e):
            args += ["--ro-bind", e, e]
    soffice = find_soffice()
    if soffice and soffice.startswith("/opt/"):
        root = "/".join(soffice.split("/")[:3])
        args += ["--ro-bind", root, root]
    return args


def bwrap_works():
    if not shutil.which("bwrap"):
        return False, "bwrap (bubblewrap) not installed"
    with tempfile.TemporaryDirectory() as d:
        try:
            r = subprocess.run(bwrap_cmd(d) + ["/usr/bin/true"], capture_output=True, timeout=20)
        except (OSError, subprocess.TimeoutExpired) as exc:
            return False, str(exc)
        if r.returncode != 0:
            return False, "bwrap cannot create namespaces here: %s" % pk.clean_text(r.stderr.decode("utf-8", "replace"), 300)
    return True, None


def check():
    ok, why = bwrap_works()
    soffice = find_soffice()
    ver = None
    if soffice and ok:
        with tempfile.TemporaryDirectory() as d:
            os.makedirs(os.path.join(d, "home"))
            os.makedirs(os.path.join(d, "tmp"))
            try:
                r = subprocess.run(bwrap_cmd(d) + [soffice, "--version"], capture_output=True, timeout=60)
                ver = pk.clean_text(r.stdout.decode("utf-8", "replace").strip(), 200)
            except (OSError, subprocess.TimeoutExpired):
                pass
    return {"ok": bool(ok and soffice), "platform": sys.platform, "soffice": soffice,
            "soffice_version": ver, "bwrap": ok, "bwrap_error": why,
            "pdftoppm": shutil.which("pdftoppm"),
            "can_render": bool(ok and soffice), "can_png": bool(ok and soffice and shutil.which("pdftoppm")),
            "error": None if ok and soffice else (why or "LibreOffice (soffice) not found")}


MEMORY_MB = 4096


def run_boxed(cmd, work, timeout):
    p = subprocess.Popen(bwrap_cmd(work) + cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                         start_new_session=True, env={"PATH": "/usr/bin:/bin"},
                         preexec_fn=limit_memory(MEMORY_MB))
    try:
        out, err = p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        os.killpg(p.pid, signal.SIGKILL)
        p.communicate()
        raise TimeoutError
    return p.returncode, out.decode("utf-8", "replace"), err.decode("utf-8", "replace")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("input", nargs="?")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--to", choices=["pdf", "png", "pptx"], default="pdf")
    ap.add_argument("--notes-pages", action="store_true", help="PDF of notes pages instead of slides")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--timeout", type=float, default=180)
    ap.add_argument("--force", action="store_true", help="replace an existing OUT file (no backup)")
    ap.add_argument("--memory-mb", type=int, default=4096, help="address-space cap for LibreOffice")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    global MEMORY_MB
    if a.memory_mb < 512:
        return pk.fail(pk.Unsupported("--memory-mb must be at least 512 (a smaller cap only makes LibreOffice fail)"), a.json)
    MEMORY_MB = a.memory_mb
    if a.check:
        res = check()
        pk.emit(res, a.json, lambda r: print(json.dumps(r, indent=2)))
        return 0 if res["ok"] else 2
    try:
        if not a.input or not a.output:
            raise pk.Unsupported("give IN and OUT, or --check")
        src = os.path.abspath(a.input)
        out = os.path.abspath(a.output)
        ext = os.path.splitext(src)[1].lower()
        if ext not in PKG_EXT | LEGACY_EXT:
            raise pk.Unsupported("unsupported input extension %s" % ext)
        if os.path.getsize(src) > MAX_INPUT:
            raise pk.Refused("input larger than %d bytes" % MAX_INPUT)
        check_input(src, ext)
        want = {"pdf": ".pdf", "pptx": ".pptx"}.get(a.to)
        if want and os.path.splitext(out)[1].lower() != want:
            raise pk.Unsupported("OUT must end in %s for --to %s" % (want, a.to))
        if os.path.exists(out) and not (a.force and a.to != "png"):
            raise pk.Unsupported("output exists: %s" % out)
        soffice = find_soffice()
        if not soffice:
            raise pk.Unsupported("LibreOffice (soffice) not found; see references/fallbacks.md")
        ok, why = bwrap_works()
        if not ok:
            raise pk.Unsupported("cannot sandbox LibreOffice: %s. Refusing to render untrusted decks "
                                 "without isolation; see references/fallbacks.md" % why)
        if a.to == "png" and not shutil.which("pdftoppm"):
            raise pk.Unsupported("pdftoppm (poppler-utils) not found")
        with tempfile.TemporaryDirectory(prefix="pptx-render-") as work:
            for d in ("home", "tmp", "out", "profile"):
                os.makedirs(os.path.join(work, d))
            os.makedirs(os.path.join(work, "profile", "user"))
            with open(os.path.join(work, "profile", "user", "registrymodifications.xcu"), "w") as fh:
                fh.write(XCU)
            name = "input" + ext
            shutil.copyfile(src, os.path.join(work, name))
            if a.to == "pptx":
                target = "pptx:Impress MS PowerPoint 2007 XML"
            elif a.notes_pages:
                target = 'pdf:impress_pdf_Export:{"ExportNotesPages":{"type":"boolean","value":"true"},' \
                         '"ExportOnlyNotesPages":{"type":"boolean","value":"true"}}'
            else:
                target = "pdf:impress_pdf_Export"
            cmd = [soffice, "-env:UserInstallation=file:///work/profile", "--headless", "--invisible",
                   "--norestore", "--nolockcheck", "--nodefault", "--convert-to", target,
                   "--outdir", "/work/out", "/work/" + name]
            try:
                rc, so, se = run_boxed(cmd, work, a.timeout)
            except TimeoutError:
                res = {"ok": False, "error": "LibreOffice did not finish within %s s; killed, nothing written" % a.timeout, "exit": 3}
                pk.emit(res, a.json, lambda r: print(r["error"]))
                return 3
            produced = os.path.join(work, "out", "input." + ("pptx" if a.to == "pptx" else "pdf"))
            if not os.path.exists(produced) or os.path.getsize(produced) == 0:
                res = {"ok": False, "error": "LibreOffice produced no output (exit %d): %s" % (rc, pk.clean_text(se, 400)), "exit": 1}
                pk.emit(res, a.json, lambda r: print(r["error"]))
                return 1
            files = []
            if a.to == "png":
                os.makedirs(os.path.join(work, "png"))
                try:
                    rc, so, se = run_boxed(["pdftoppm", "-png", "-r", str(a.dpi), "/work/out/input.pdf",
                                            "/work/png/slide"], work, a.timeout)
                except TimeoutError:
                    res = {"ok": False, "error": "pdftoppm timed out; nothing written", "exit": 3}
                    pk.emit(res, a.json, lambda r: print(r["error"]))
                    return 3
                pngs = sorted(os.listdir(os.path.join(work, "png")))
                if not pngs:
                    res = {"ok": False, "error": "pdftoppm produced no images", "exit": 1}
                    pk.emit(res, a.json, lambda r: print(r["error"]))
                    return 1
                staging = tempfile.mkdtemp(prefix=".png-", dir=os.path.dirname(out) or ".")
                for f in pngs:
                    shutil.move(os.path.join(work, "png", f), os.path.join(staging, f))
                try:
                    os.rename(staging, out)  # fails if out appeared meanwhile (non-empty dir or file)
                except OSError:
                    shutil.rmtree(staging)
                    raise pk.Unsupported("output appeared while rendering: %s" % out)
                files = [os.path.join(out, f) for f in pngs]
            else:
                tmp = tempfile.NamedTemporaryFile(delete=False, dir=os.path.dirname(out) or ".", prefix=".render-")
                tmp.close()
                shutil.copyfile(produced, tmp.name)
                if a.force and os.path.exists(out):
                    os.replace(tmp.name, out)
                else:
                    try:
                        os.link(tmp.name, out)
                    finally:
                        os.unlink(tmp.name)
                files = [out]
        res = {"ok": True, "input": a.input, "output": a.output, "to": a.to, "files": files,
               "engine": "LibreOffice (sandboxed bwrap, no network)", "error": None,
               "note": "Rendered by LibreOffice, not PowerPoint; layout and fonts may differ."}
        if a.to == "pptx":
            res["note"] = ("Converted by LibreOffice: the whole deck was re-saved and keeps any external "
                           "links, notes and metadata. Run inspect_pptx.py and pptx_sanitize.py --check-clean on it.")
        pk.emit(res, a.json, lambda r: print("wrote %d file(s): %s" % (len(r["files"]), ", ".join(r["files"][:5]))))
        return 0
    except (pk.Unsupported, pk.Refused) as exc:
        return pk.fail(exc, a.json)


if __name__ == "__main__":
    sys.exit(main())
