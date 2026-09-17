"""Resolve a manuscript's numbered reference list to Zotero items.

    python resolve.py --docx manuscript.docx --out-dir ./work

Writes to --out-dir:
    refmap.json     {reference number: Zotero item key}
    itemmeta.json   {item key: {itemID, csl}}   (feeds inject.py)
    citation-map.md human-readable audit table
    unresolved.md   references that could not be matched (empty file if none)

Nothing is written to the user's Zotero library. If references are missing, the
caller decides whether to import them.
"""
from __future__ import annotations

import argparse
import json
import pathlib

import zcommon as Z


def match(entry, library):
    best, best_key, best_title = 0.0, None, None
    for it in library:
        title = it["data"].get("title") or ""
        if len(Z.tokens(title)) < 4:
            continue
        score = Z.containment(title, entry)
        if score > best:
            best, best_key, best_title = score, it["key"], title
    return best, best_key, best_title


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--threshold", type=float, default=0.85,
                    help="token-containment score to accept a title match")
    args = ap.parse_args()

    out = pathlib.Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)

    Z.check_zotero_running()
    refs = Z.references(args.docx)
    if not refs:
        raise SystemExit("Parsed no numbered reference entries — check the list format.")
    library = Z.library_items()
    print(f"references: {len(refs)}   library items: {len(library)}")

    refmap, how, unresolved = {}, {}, []
    for n in sorted(refs):
        entry = refs[n]
        score, key, title = match(entry, library)
        if score >= args.threshold:
            refmap[n], how[n] = key, f"title match {score:.2f}"
            continue

        # Second pass: the manuscript's wording often differs from the Zotero
        # title, so go via the DOI to get an authoritative title and match that.
        doi = Z.extract_doi(entry)
        resolved_key, resolved_score = None, 0.0
        if doi:
            auth = Z.crossref_title(doi)
            if auth:
                for it in library:
                    t = it["data"].get("title") or ""
                    s = Z.containment(t, auth)
                    if s > resolved_score:
                        resolved_score, resolved_key = s, it["key"]
        if resolved_key and resolved_score >= args.threshold:
            refmap[n], how[n] = resolved_key, f"DOI→CrossRef→library {resolved_score:.2f}"
        else:
            unresolved.append((n, doi, entry, score))

    # metadata for every cited item
    ids = Z.item_ids()
    meta = {}
    for key in sorted(set(refmap.values())):
        meta[key] = {"itemID": ids.get(key), "csl": Z.csl(key)}
        if meta[key]["itemID"] is None:
            print(f"  warning: no numeric itemID for {key} (not in zotero.sqlite?)")

    (out / "refmap.json").write_text(json.dumps(refmap, indent=1), encoding="utf-8")
    (out / "itemmeta.json").write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")

    lines = ["# 参考文献 → Zotero 条目", "",
             f"- 文稿：`{args.docx}`",
             f"- 参考文献条目：{len(refs)}",
             f"- 已定位：**{len(refmap)}**　未定位：**{len(unresolved)}**", "",
             "| # | Zotero Key | 匹配依据 |", "|---|---|---|"]
    for n in sorted(refs):
        lines.append(f"| {n} | `{refmap[n]}` | {how[n]} |" if n in refmap else f"| {n} | — | **未定位** |")
    (out / "citation-map.md").write_text("\n".join(lines), encoding="utf-8")

    ulines = ["# 未定位的参考文献", "",
              "这些条目不在 Zotero 库中（标题与 DOI 两种匹配都失败）。",
              "**未经用户确认，不要导入它们的库。**", ""]
    for n, doi, entry, score in unresolved:
        ulines += [f"## [{n}]", f"- DOI: `{doi or '未找到'}`", f"- 最佳标题匹配度: {score:.2f}",
                   f"- 原文: {entry[:200]}", ""]
    (out / "unresolved.md").write_text("\n".join(ulines), encoding="utf-8")

    print(f"resolved {len(refmap)}/{len(refs)}")
    if unresolved:
        print(f"UNRESOLVED ({len(unresolved)}) — report these to the user, do not import:")
        for n, doi, entry, score in unresolved:
            print(f"  [{n}] {doi or '(no DOI)'}  {entry[:70]}")
    print(f"wrote refmap.json, itemmeta.json, citation-map.md, unresolved.md in {out}")


if __name__ == "__main__":
    main()
