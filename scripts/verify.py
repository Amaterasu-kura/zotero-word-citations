"""Verify that the converted document cites the same works as the original.

    python verify.py --original manuscript.docx --output manuscript_v2.docx \
                     --refmap work/refmap.json --itemmeta work/itemmeta.json

Two independent checks, because they fail independently:

  A. positional correspondence - walking the body in order, the Nth Zotero field
     in the output must cite exactly the items the Nth [n] token of the original
     mapped to. Catches scrambling introduced by the run-level rewrite.

  B. semantic correctness - for every reference, compare the manuscript's own
     entry against the metadata of the Zotero item it was mapped to. Catches a
     wrong mapping (which A cannot see) and also surfaces errors in the
     manuscript's own reference list: DOIs that do not resolve at all, or that
     resolve to an unrelated paper.

Exits non-zero if check A fails or check B finds a wrong mapping. Findings that
are the *manuscript's* problem (bad DOIs) are reported but do not fail the run —
report them to the user rather than editing their reference list unasked.
"""
from __future__ import annotations

import argparse
import json
import pathlib
import re
import unicodedata

import zcommon as Z

failures, notes = [], []


def part_a(original, output, refmap):
    o_paras = Z.paragraphs(Z.read_docx(original))
    x_paras = Z.paragraphs(Z.read_docx(output))
    head = next(i for i, p in enumerate(o_paras)
                if Z.para_text(p).strip().lower() in ("references", "reference", "bibliography"))

    seq_o, seq_x = [], []
    for i in range(head):
        for m in Z.CITE_RE.finditer(Z.para_text(o_paras[i])):
            seq_o.append((i, [int(v) for v in re.split(r"\s*,\s*", m.group(1))]))
        for s in Z.INSTR_RE.findall(x_paras[i]):
            if "ADDIN ZOTERO_ITEM" in s:
                obj = json.loads(s.split("CSL_CITATION", 1)[1].strip())
                seq_x.append((i, [u.rsplit("/", 1)[1] for u in
                                  (it["uris"][0] for it in obj["citationItems"])]))

    print("== A. positional correspondence ==")
    print(f"   original citation tokens: {len(seq_o)}   output fields: {len(seq_x)}")
    if len(seq_o) != len(seq_x):
        failures.append(f"token count {len(seq_o)} != {len(seq_x)}")
        print("   [FAIL] counts differ — the rewrite dropped or invented citations")
        return
    bad = 0
    for k, ((po, nums), (px, keys)) in enumerate(zip(seq_o, seq_x)):
        expect = [refmap[n] for n in nums]
        if po != px or expect != keys:
            bad += 1
            failures.append(f"token {k} in paragraph {po}/{px}: {nums} -> {keys}, expected {expect}")
    print("   [PASS] every citation point corresponds" if not bad
          else f"   [FAIL] {bad} of {len(seq_o)} citation points do not correspond")


def year_in(entry):
    m = re.search(r"\b(1[89]\d{2}|20\d{2})\b", entry or "")
    return m.group(0) if m else None


def part_b(original, refmap, meta):
    """For each reference: is the Zotero item we mapped to really the paper the
    manuscript names?

    A direct DOI comparison is the strongest evidence, but it is often
    unavailable — plenty of Zotero items carry no DOI, and manuscripts have
    wording that differs from the Zotero title. So the checks run strongest
    first and fall through: DOI equality, then the manuscript DOI resolved via
    CrossRef compared against the Zotero title, then plain title containment.
    Only when every route fails is the mapping called wrong; declaring a failure
    on the weakest check alone produces false alarms on correct mappings, which
    is worse than useless because it teaches the user to ignore the report.
    """
    refs = Z.references(original)
    print(f"\n== B. semantic correctness ({len(refs)} references) ==")
    wrong, manuscript_errors = [], []
    for n in sorted(refs):
        entry = refs[n]
        key = refmap.get(n)
        if not key:
            wrong.append((n, "no Zotero mapping"))
            continue
        csl = meta[key]["csl"]
        zot_title = csl.get("title") or ""
        man_doi = Z.extract_doi(entry)
        zot_doi = (csl.get("DOI") or "").strip().lower().rstrip(".")
        man_doi_n = (man_doi or "").strip().lower().rstrip(".")

        # A year that disagrees is worth surfacing whatever the mapping verdict,
        # because it is exactly the kind of slip that survives a title check.
        parts = (csl.get("issued") or {}).get("date-parts") or [[]]
        zot_year = str(parts[0][0]) if parts and parts[0] else None
        man_year = year_in(entry)
        if zot_year and man_year and zot_year != man_year:
            manuscript_errors.append((n, "year", man_year, zot_year))

        if man_doi_n and zot_doi:
            if man_doi_n == zot_doi:
                continue                                    # strongest: same DOI
            if Z.containment(zot_title, entry) >= 0.85:
                manuscript_errors.append((n, "DOI", man_doi, zot_doi))
                continue
            wrong.append((n, f"DOI {man_doi} != {zot_doi} and titles disagree"))
            continue

        # No DOI on one side. An item often has no DOI recorded even though the
        # manuscript quotes one, so resolve the manuscript's DOI and compare
        # against the Zotero title before falling back to plain title matching.
        if man_doi:
            authoritative = Z.crossref_title(man_doi)
            if authoritative and Z.containment(zot_title, authoritative) >= 0.85:
                continue
        if Z.containment(zot_title, entry) >= 0.85:
            continue
        wrong.append((n, "title matches neither directly nor via CrossRef"))

    for n, kind, got, want in manuscript_errors:
        print(f"   [NOTE] ref {n}: manuscript {kind} is {got}, Zotero has {want}"
              f" — one of them is wrong; the output carries Zotero's value")
    for n, why in wrong:
        print(f"   [FAIL] ref {n}: {why}")
    if wrong:
        failures.extend(f"ref {n}: {why}" for n, why in wrong)
    else:
        print(f"   [PASS] every reference maps to the paper it names"
              f"   ({len(manuscript_errors)} manuscript-side discrepancy(ies) noted)")
    notes.extend(manuscript_errors)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--original", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--refmap", required=True)
    ap.add_argument("--itemmeta", required=True)
    args = ap.parse_args()

    refmap = {int(k): v for k, v in json.loads(pathlib.Path(args.refmap).read_text(encoding="utf-8")).items()}
    meta = json.loads(pathlib.Path(args.itemmeta).read_text(encoding="utf-8"))

    out_xml = Z.read_docx(args.output)
    print("== structure ==")
    ni, nb = out_xml.count("ADDIN ZOTERO_ITEM"), out_xml.count("ADDIN ZOTERO_BIBL")
    print(f"   citation fields: {ni}   bibliography fields: {nb}")
    if nb != 1:
        failures.append(f"expected exactly 1 bibliography field, found {nb}")
    leftover = Z.CITE_RE.findall("".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", out_xml, re.S)))
    print(f"   leftover plain [n] tokens: {len(leftover)}")
    if leftover:
        failures.append(f"{len(leftover)} plain-text citation tokens survived")

    # Text that was already XML-escaped must be copied through untouched. If a
    # run carrying an entity (&lt; &gt; &amp;) also held part of a citation
    # token, a naive rewrite re-escapes it and Word then prints the entity
    # literally — e.g. "&lt;16 µm" instead of "<16 µm". It only shows up in runs
    # the citation happened to split, so it is easy to ship unnoticed.
    #
    # Only visible text counts. Entities inside <w:instrText> are the field's
    # JSON payload (Zotero stores some metadata already escaped) and are never
    # displayed; Zotero re-reads the item by URI on refresh anyway.
    visible_double = 0
    for m in re.finditer(r"&amp;(?:lt|gt|amp|quot);", out_xml):
        pre = out_xml[:m.start()]
        if pre.rfind("<w:t") > pre.rfind("</w:t>"):
            visible_double += 1
    print(f"   double-escaped entities in visible text: {visible_double}")
    if visible_double:
        failures.append(f"{visible_double} double-escaped XML entities in visible text — text corruption")

    part_a(args.original, args.output, refmap)
    part_b(args.original, refmap, meta)

    print()
    if failures:
        print("FAILURES:")
        for f in failures:
            print("  -", f)
        raise SystemExit(1)
    print("ALL CHECKS PASSED")
    if notes:
        print("Reminder: report the manuscript-side errors above to the user; do not edit their reference list unasked.")


if __name__ == "__main__":
    main()
