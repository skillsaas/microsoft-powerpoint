"""Hyperlink checks for the microsoft-powerpoint scripts (not a public API).

Written for PresentationML; the rules follow lessons from the microsoft-word skill's review
rounds, re-implemented here rather than copied.

classify(target) -> scheme, host, risky flag
    Risky: any scheme other than http, https and mailto (javascript:, vbscript:, data:, file:,
    smb:, ms-*, custom handlers), UNC paths (\\\\host\\share, which Windows may try to log in to),
    drive paths (C:\\...) and relative paths (they open a file next to the deck).

mismatch(text, target) -> reason or None
    The words a viewer sees on a link point somewhere else:
      "userinfo"          the text shows a host before "@" (https://bank.com@evil.io), unless the
                          target is mailto:
      "no_web_host"       the text shows a domain but the target is not an http(s) or mailto URL
      "different_domain"  the first domain in the text is not the target's site
      "invisible_chars"   the text contains zero-width, bidi or other format characters
    Text is NFKC-normalized, full-width and ideographic dots count as dots, and hosts are
    compared in IDNA form without a trailing dot. "Site" uses a short built-in list of
    multi-label suffixes (co.uk, gov.pk, github.io ...), not the full Public Suffix List; a
    suffix missing from the list makes two sites under it look like one (a missed flag, never
    a false one).
"""
import re
import unicodedata
from urllib.parse import urlsplit

SAFE_SCHEMES = {"http", "https", "mailto"}
_DOTS = str.maketrans({"\u3002": ".", "\uff0e": ".", "\uff61": "."})
_INVISIBLE = re.compile("[\u00ad\u061c\u180e\u200b-\u200f\u202a-\u202e\u2060-\u2064\u2066-\u2069\ufeff]")
_DOMAIN = re.compile(r"(?<![\w@.-])((?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z][a-z0-9-]{1,62})\.?(?![\w-])",
                     re.I)
_USERINFO = re.compile(r"(?:(https?|ftp)://)?([^\s/@]+)@([a-z0-9.-]+\.[a-z]{2,})", re.I)
# Text that looks like a file name, not a site ("report.pptx", "logo.png").
_FILE_TLDS = {"ppt", "pptx", "pptm", "ppsx", "potx", "doc", "docx", "xls", "xlsx", "pdf", "txt", "csv",
              "png", "jpg", "jpeg", "gif", "svg", "zip", "md", "json", "xml", "htm", "html", "py", "js", "mp4"}
SUFFIXES = {
    "co.uk", "org.uk", "ac.uk", "gov.uk", "me.uk", "com.au", "net.au", "org.au", "gov.au", "edu.au",
    "co.nz", "org.nz", "govt.nz", "ac.nz", "co.jp", "ne.jp", "or.jp", "ac.jp", "go.jp", "co.kr", "go.kr",
    "ac.kr", "com.br", "gov.br", "com.cn", "gov.cn", "edu.cn", "com.hk", "com.sg", "gov.sg", "com.tw",
    "co.in", "gov.in", "ac.in", "nic.in", "com.pk", "org.pk", "net.pk", "gov.pk", "edu.pk", "gob.pk",
    "web.pk", "com.mx", "gob.mx", "com.ar", "com.tr", "gov.tr", "co.za", "gov.za", "ac.za", "co.il",
    "ac.il", "gov.il", "com.sa", "gov.sa", "gov.ae", "com.eg", "gov.eg", "co.id", "go.id", "com.my",
    "gov.my", "com.ph", "gov.ph", "com.vn", "com.ng", "gov.ng", "co.ke", "go.ke", "com.bd", "gov.bd",
    "gc.ca", "gouv.fr",
    "github.io", "gitlab.io", "pages.dev", "workers.dev", "netlify.app", "vercel.app", "web.app",
    "firebaseapp.com", "appspot.com", "herokuapp.com", "azurewebsites.net", "cloudfront.net",
    "blogspot.com", "wordpress.com", "ngrok.io", "ngrok-free.app", "onrender.com", "fly.dev",
}


def norm_host(h):
    h = unicodedata.normalize("NFKC", h or "").translate(_DOTS).strip().rstrip(".").lower()
    out = []
    for label in h.split("."):
        try:
            out.append(label.encode("idna").decode("ascii") if label and not label.isascii() else label)
        except UnicodeError:
            out.append(label)
    return ".".join(out)


def site(h):
    labels = norm_host(h).split(".")
    n = 3 if len(labels) >= 3 and ".".join(labels[-2:]) in SUFFIXES else 2
    return ".".join(labels[-n:])


def same_site(a, b):
    a, b = norm_host(a), norm_host(b)
    if a in SUFFIXES or b in SUFFIXES:
        return a == b
    return site(a) == site(b)


def classify(target):
    """(scheme, host, risky) for a relationship target or URL."""
    t = (target or "").strip()
    if t.startswith(("\\\\", "//")) and not t.startswith("///"):
        host = t.lstrip("\\/").replace("\\", "/").split("/", 1)[0]
        return "unc", norm_host(host), True
    if re.match(r"^[a-zA-Z]:[\\/]", t):
        return "file", "", True
    try:
        sp = urlsplit(t)
    except ValueError:
        return "invalid", "", True
    scheme = (sp.scheme or "").lower()
    if not scheme:
        return "relative", "", True  # an external relative target opens a file next to the deck
    host = ""
    if scheme == "mailto":
        host = norm_host(sp.path.split("?", 1)[0].rsplit("@", 1)[-1])
    else:
        try:
            host = norm_host(sp.hostname or "")
        except ValueError:
            host = ""
    return scheme, host, scheme not in SAFE_SCHEMES


def mismatch(text, target):
    raw = text or ""
    if _INVISIBLE.search(raw) or any(unicodedata.category(c) == "Cf" for c in raw):
        return "invisible_chars"
    t = unicodedata.normalize("NFKC", raw).translate(_DOTS)
    scheme, host, _ = classify(target)
    if scheme != "mailto":
        for m in _USERINFO.finditer(t):
            user, after = m.group(2), m.group(3)
            looks_host = "." in user and user.rsplit(".", 1)[-1].lower() not in _FILE_TLDS
            if m.group(1) or (looks_host and not (host and same_site(after, host))):
                return "userinfo"
    shown = [d for d in _DOMAIN.findall(t) if d.rsplit(".", 1)[-1].lower() not in _FILE_TLDS]
    if not shown:
        return None
    if scheme not in SAFE_SCHEMES or not host:
        return "no_web_host"
    if scheme == "mailto":
        return None if same_site(shown[-1], host) or same_site(shown[0], host) else "different_domain"
    if not same_site(shown[0], host):
        return "different_domain"
    return None
