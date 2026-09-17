"""Inject live Zotero citation fields into a Word manuscript.

    python inject.py --docx manuscript.docx --refmap work/refmap.json \
                     --itemmeta work/itemmeta.json \
                     --style http://www.zotero.org/styles/elsevier-harvard \
                     --out manuscript_v2.docx

The output is a rewrite of the original package: every other part is copied
through untouched. Never rebuild a docx from scratch for this — a hand-built
minimal document puts Word into compatibility mode, and Zotero's refresh then
freezes Word with no error at all. See references/field-format.md.

The body pass is a run-level rewrite rather than a find/replace because citation
tokens are routinely split across several <w:t> runs (red colour marks,
spell-check boundaries, revision marks cut through them). A regex over individual
run texts silently misses those, which is a failure you cannot see in the output.
So each paragraph is joined into one string, tokens are located there, and the run
sequence is rebuilt around them.
"""
from __future__ import annotations

import argparse
import json
import pathlib
import random
import re
import string
import zipfile
from xml.sax.saxutils import escape

import zcommon as Z

SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
BIBL_INSTR = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '


def author_date(csl):
    """Approximate author-date text for the field result, so the document still
    reads correctly before the user refreshes."""
    authors = csl.get("author") or []
    parts = (csl.get("issued") or {}).get("date-parts") or [[]]
    year = parts[0][0] if parts and parts[0] else "n.d."
    fam = [a.get("family") or a.get("literal") or "" for a in authors]
    if not fam:
        name = (csl.get("title") or "")[:25]
    elif len(fam) == 1:
        name = fam[0]
    elif len(fam) == 2:
        name = f"{fam[0]} and {fam[1]}"
    else:
        name = f"{fam[0]} et al."
    return f"{name}, {year}"


def citation_field(refnums, refmap, meta, rpr=""):
    items, shown = [], []
    for n in refnums:
        m = meta[refmap[n]]
        items.append({"id": m["itemID"], "uris": [m["csl"]["id"]], "itemData": m["csl"]})
        shown.append(author_date(m["csl"]))
    text = "; ".join(shown)
    payload = {
        "citationID": "".join(random.choices(string.ascii_letters + string.digits, k=8)),
        "properties": {"unsorted": False, "formattedCitation": f"({text})",
                       "plainCitation": f"({text})", "noteIndex": 0},
        "citationItems": items,
        "schema": SCHEMA,
    }
    instr = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, separators=(",", ":")) + " "
    return (
        f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr}<w:instrText xml:space="preserve">{escape(instr)}</w:instrText></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:t xml:space="preserve">({escape(text)})</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
    )


def bibliography_field():
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve">{escape(BIBL_INSTR)}</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t xml:space="preserve">(Zotero will build the reference list on Refresh)</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


PROPERTY_RE = re.compile(r"<property\b[^>]*>.*?</property>|<property\b[^>]*/>", re.S)
PROP_NS = ('xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"')
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


def build_prefs(style_id, existing_xml=None):
    """Build docProps/custom.xml, keeping any custom properties already present.

    Zotero reads the citation style from custom document properties, not from the
    fields, and Office caps each property value at 255 characters — hence the
    chunking across ZOTERO_PREF_1, _2, …

    Merging rather than replacing matters: documents carry unrelated custom
    properties (authoring-tool markers, workflow metadata) and silently dropping
    them is data loss the user never asked for and would not see.
    """
    data = (
        '<data data-version="3" zotero-version="6.0.7">'
        f'<session id="{"".join(random.choices(string.ascii_letters + string.digits, k=8))}"/>'
        f'<style id="{style_id}" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
        '<prefs><pref name="fieldType" value="Field"/>'
        '<pref name="automaticJournalAbbreviations" value="true"/></prefs></data>'
    )
    chunks = [data[i:i + 255] for i in range(0, len(data), 255)]
    zprops = [
        f'<property fmtid="{FMTID}" pid="{{pid}}" name="ZOTERO_PREF_{i + 1}">'
        f"<vt:lpwstr>{escape(c)}</vt:lpwstr></property>"
        for i, c in enumerate(chunks)
    ]

    kept = []
    if existing_xml:
        for m in PROPERTY_RE.finditer(existing_xml):
            block = m.group(0)
            name = re.search(r'name="([^"]*)"', block)
            # drop stale ZOTERO_PREF_* from a previous run; keep everything else
            if name and name.group(1).startswith("ZOTERO_PREF_"):
                continue
            kept.append(re.sub(r'\spid="[^"]*"', "", block))

    props = []
    for i, block in enumerate(kept + zprops):
        props.append(block.replace("{pid}", str(i + 2), 1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<Properties {PROP_NS}>{''.join(props)}</Properties>")


def parse_runs(paragraph):
    out = []
    for m in Z.RUN_RE.finditer(paragraph):
        run = m.group(0)
        tm = Z.TEXT_RE.search(run)
        if not tm:
            continue
        rpr = Z.RPR_RE.search(run)
        out.append({
            "p0": m.start(), "p1": m.end(),
            "xml": run, "text": tm.group(2),
            "rpr": strip_red(rpr.group(0)) if rpr else "",
        })
    return out


def strip_red(xml):
    return re.sub(r'<w:color\s+w:val="FF0000"\s*/>', "", xml)


def replace_run_text(run_xml, text):
    """Swap a run's <w:t> content.

    `text` is taken verbatim from the original run, so it is *already* in XML
    escaped form (`&lt;16 µm`, `&gt;400`, `&amp;`). Re-escaping it here would
    turn those into `&amp;lt;16 µm` and Word would render the entity literally —
    corruption that is easy to miss because it only hits runs that carry both an
    entity and part of a citation token. Insert it unchanged.
    """
    tm = Z.TEXT_RE.search(run_xml)
    return run_xml[:tm.start(2)] + text + run_xml[tm.end(2):]


def inject_paragraph(paragraph, refmap, meta):
    runs = parse_runs(paragraph)
    if not runs:
        return paragraph, 0
    joined = "".join(r["text"] for r in runs)
    toks = []
    for m in Z.CITE_RE.finditer(joined):
        nums = [int(x) for x in re.split(r"\s*,\s*", m.group(1))]
        if all(n in refmap for n in nums):
            toks.append((m.start(), m.end(), nums))
    if not toks:
        return paragraph, 0

    offsets, pos = {}, 0
    for r in runs:
        offsets[id(r)] = pos
        pos += len(r["text"])

    pieces, emitted = [], 0
    for r in runs:
        rs = offsets[id(r)]
        hits = [t for t in toks if t[0] < rs + len(r["text"]) and t[1] > rs]
        if not hits:
            pieces.append((r["p0"], r["p1"], strip_red(r["xml"])))
            continue
        chunks, cursor = [], 0
        for ts, te, nums in sorted(hits, key=lambda t: t[0]):
            ls, le = max(0, ts - rs), min(len(r["text"]), te - rs)
            if ls > cursor:
                chunks.append(("text", r["text"][cursor:ls]))
            if ts >= rs:                      # token begins in this run -> field here
                chunks.append(("field", nums))
                emitted += 1
            cursor = le
        if cursor < len(r["text"]):
            chunks.append(("text", r["text"][cursor:]))
        built = []
        for kind, val in chunks:
            if kind == "field":
                built.append(citation_field(val, refmap, meta, r["rpr"]))
            elif val:
                built.append(replace_run_text(r["xml"], val))
        pieces.append((r["p0"], r["p1"], "".join(built)))

    out, last = [], 0
    for p0, p1, xml in pieces:
        out.append(paragraph[last:p0])
        out.append(xml)
        last = p1
    out.append(paragraph[last:])
    return "".join(out), emitted


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True)
    ap.add_argument("--refmap", required=True)
    ap.add_argument("--itemmeta", required=True)
    ap.add_argument("--style", required=True,
                    help="CSL style URL, e.g. http://www.zotero.org/styles/elsevier-harvard")
    ap.add_argument("--out", required=True)
    ap.add_argument("--keep-red", action="store_true",
                    help="do not strip direct red character colour")
    args = ap.parse_args()

    refmap = {int(k): v for k, v in json.loads(pathlib.Path(args.refmap).read_text(encoding="utf-8")).items()}
    meta = json.loads(pathlib.Path(args.itemmeta).read_text(encoding="utf-8"))

    if args.out == args.docx:
        raise SystemExit("Refusing to overwrite the input document — pass a different --out.")

    zin = zipfile.ZipFile(args.docx)
    doc = zin.read("word/document.xml").decode("utf-8")
    if not args.keep_red:
        doc = strip_red(doc)

    paras = [(m.start(), m.end(), m.group(0)) for m in Z.PARA_RE.finditer(doc)]
    texts = [Z.para_text(p) for _, _, p in paras]
    try:
        ref_idx = next(i for i, t in enumerate(texts) if t.strip().lower() in ("references", "reference", "bibliography"))
    except StopIteration:
        raise SystemExit("No 'References' heading found in the document body.")

    total, edits = 0, []
    for i in range(ref_idx):
        new, n = inject_paragraph(paras[i][2], refmap, meta)
        if n:
            edits.append((paras[i][0], paras[i][1], new))
            total += n
    for p0, p1, new in reversed(edits):
        doc = doc[:p0] + new + doc[p1:]
    print(f"citation fields injected: {total}")

    # swap the plain reference list for the bibliography field
    paras = [(m.start(), m.end(), m.group(0)) for m in Z.PARA_RE.finditer(doc)]
    texts = [Z.para_text(p) for _, _, p in paras]
    ref_idx = next(i for i, t in enumerate(texts) if t.strip().lower() in ("references", "reference", "bibliography"))
    head_end = paras[ref_idx][1]
    sect = doc.rindex("<w:sectPr")
    removed = sect - head_end
    doc = doc[:head_end] + f"<w:p>{bibliography_field()}</w:p>" + doc[sect:]
    print(f"bibliography: replaced {removed} chars of the plain list with the ZOTERO_BIBL field")

    custom = build_prefs(args.style)
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    if "custom-properties" not in ct:
        ct = ct.replace("</Types>",
                        '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
    rels = zin.read("_rels/.rels").decode("utf-8")
    if "custom-properties" not in rels:
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>')

    with zipfile.ZipFile(args.out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in zin.namelist():
            if name == "word/document.xml":
                zo.writestr(name, doc)
            elif name == "[Content_Types].xml":
                zo.writestr(name, ct)
            elif name == "_rels/.rels":
                zo.writestr(name, rels)
            elif name == "docProps/custom.xml":
                continue
            else:
                zo.writestr(name, zin.read(name))
        zo.writestr("docProps/custom.xml", custom)

    print(f"wrote {args.out}")
    print(f"  ZOTERO_ITEM fields: {doc.count('ADDIN ZOTERO_ITEM')}")
    print(f"  ZOTERO_BIBL fields: {doc.count('ADDIN ZOTERO_BIBL')}")
    print(f"  style: {args.style}")


if __name__ == "__main__":
    main()
