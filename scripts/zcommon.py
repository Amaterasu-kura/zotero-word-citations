"""Shared helpers for the zotero-word-citations scripts.

Stdlib only, so these run under any Python 3.9+ without a virtualenv.
"""
from __future__ import annotations

import json
import os
import pathlib
import re
import shutil
import sqlite3
import tempfile
import urllib.parse
import urllib.request
import zipfile

API = "http://127.0.0.1:23119/api/users/0"

PARA_RE = re.compile(r"<w:p(?:\s[^>]*)?>.*?</w:p>|<w:p(?:\s[^>]*)?/>", re.S)
RUN_RE = re.compile(r"<w:r(?:\s[^>]*)?>.*?</w:r>|<w:r(?:\s[^>]*)?/>", re.S)
TEXT_RE = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)
RPR_RE = re.compile(r"<w:rPr>.*?</w:rPr>", re.S)
INSTR_RE = re.compile(r"<w:instrText[^>]*>(.*?)</w:instrText>", re.S)
CITE_RE = re.compile(r"\[(\d{1,3}(?:\s*,\s*\d{1,3})*)\]")

STOP = {"the", "a", "an", "of", "in", "and", "for", "on", "to", "with", "at", "by"}


# --------------------------------------------------------------------- Zotero

def _profile_prefs():
    """prefs.js files of every Zotero profile on this machine."""
    out = []
    if os.name == "nt":
        root = pathlib.Path(os.environ.get("APPDATA", "")) / "Zotero" / "Zotero" / "Profiles"
    elif os.uname().sysname == "Darwin":
        root = pathlib.Path.home() / "Library" / "Application Support" / "Zotero" / "Profiles"
    else:
        root = pathlib.Path.home() / ".zotero" / "zotero"
    if root.is_dir():
        out.extend(sorted(root.glob("*/prefs.js")))
    return out


def zotero_data_dir():
    """Where Zotero keeps zotero.sqlite, storage/, styles/ — read from prefs.js.

    Users move this; the default ~/Zotero is often a stale leftover that still
    exists, so trusting it silently indexes the wrong library.
    """
    for prefs in _profile_prefs():
        try:
            text = prefs.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        m = re.search(r'"extensions\.zotero\.dataDir"\s*,\s*"((?:[^"\\]|\\.)*)"', text)
        if m:
            return pathlib.Path(m.group(1).replace("\\\\", "\\"))
    return pathlib.Path.home() / "Zotero"


def item_ids():
    """{item key: numeric itemID} from a WAL-inclusive snapshot of zotero.sqlite.

    Zotero holds an exclusive lock on the live database, so the file plus its
    -wal and -shm siblings are copied before opening.
    """
    data = zotero_data_dir()
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="zot_snap_"))
    try:
        for name in ("zotero.sqlite", "zotero.sqlite-wal", "zotero.sqlite-shm"):
            src = data / name
            if src.exists():
                shutil.copy2(src, tmp / name)
        db = tmp / "zotero.sqlite"
        if not db.exists():
            raise SystemExit(f"zotero.sqlite not found under {data}")
        con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
        return {key: iid for iid, key in con.execute("SELECT itemID, key FROM items")}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def api(path, timeout=40):
    req = urllib.request.Request(API + path, headers={"User-Agent": "zotero-word-citations"})
    return json.load(urllib.request.urlopen(req, timeout=timeout))


def library_items():
    return api("/items/top?format=json&limit=1000")


def csl(key):
    data = api(f"/items/{key}?format=csljson")
    return data[0] if isinstance(data, list) else data


def check_zotero_running():
    try:
        api("/items?limit=1", timeout=8)
    except Exception as e:
        raise SystemExit(
            "Zotero's local API is not answering on 127.0.0.1:23119.\n"
            "Start Zotero, and enable Settings → Advanced → "
            '"Allow other applications on this computer to communicate with Zotero".\n'
            f"({type(e).__name__}: {e})"
        )


# ---------------------------------------------------------------------- text

def fold(text):
    return re.sub(r"[^a-z0-9]+", "", (text or "").lower())


def tokens(text):
    return set(re.findall(r"[a-z0-9]+", (text or "").lower())) - STOP


def containment(needle, haystack):
    """Fraction of `needle`'s tokens that appear in `haystack`."""
    n = tokens(needle)
    return len(n & tokens(haystack)) / len(n) if n else 0.0


def crossref_title(doi, timeout=30):
    """Authoritative title for a DOI, or None if CrossRef cannot resolve it."""
    try:
        req = urllib.request.Request(
            "https://api.crossref.org/works/" + urllib.parse.quote(doi),
            headers={"User-Agent": "zotero-word-citations"})
        data = json.load(urllib.request.urlopen(req, timeout=timeout))
        return (data["message"].get("title") or [""])[0] or None
    except Exception:
        return None


def extract_doi(text):
    m = re.search(r"doi\.org/(\S+)", text or "")
    return m.group(1).rstrip(".").rstrip(")") if m else None


# ---------------------------------------------------------------------- docx

def read_docx(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8", "replace")


def paragraphs(xml):
    return PARA_RE.findall(xml)


def para_text(paragraph):
    return "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", paragraph, re.S))


def references(path):
    """{number: entry text} for the numbered reference list at the end."""
    paras = paragraphs(read_docx(path))
    texts = [para_text(p) for p in paras]
    try:
        start = next(i for i, t in enumerate(texts) if t.strip().lower() in ("references", "reference", "bibliography"))
    except StopIteration:
        raise SystemExit("No 'References' heading found — is this the right document?")
    refs = {}
    for t in texts[start:]:
        m = re.match(r"^\s*(\d{1,3})[.)]\s+(.{25,})$", t.strip())
        if m:
            refs[int(m.group(1))] = m.group(2)
    return refs
