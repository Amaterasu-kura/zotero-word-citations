---
name: zotero-word-citations
description: Turn a Word manuscript whose citations are typed text into one with live, updatable Zotero citation fields, and regenerate its reference list from Zotero. Use this whenever a .docx has citations or a bibliography that are plain text — numbered like "[1,2]", or a reference list that cannot be refreshed — and the user wants them managed by Zotero. Triggers on requests such as "把文献插进 Word", "引用无法更新", "用 Zotero 插入参考文献", "参考文献表是纯文本", "EndNote 转 Zotero", "把这个稿子的引用变成 Zotero 的", "换引用样式", or any mention of inserting/refreshing citations in a Word document. Also use when the user wants a citation style (Elsevier, ACS, GB/T 7714, Vancouver…) applied to a manuscript, or when a reference list must be regenerated from a Zotero library.
---

# Zotero citations in Word

Convert a manuscript whose citations are typed text into one whose citations are **live Zotero fields**. After the conversion the user can click Refresh in Word to renumber everything, and can switch citation style with one click — the two things a plain-text reference list can never do.

This works by writing Zotero's own field code into the .docx. Zotero then adopts those fields on the next refresh. It is a well-defined format, but it has traps that cost real debugging time; this skill exists mainly to carry those forward.

## What the user gets

A new .docx (never overwrite the original) containing:

- one `ADDIN ZOTERO_ITEM CSL_CITATION` field per in-text citation, each carrying the Zotero item URIs it cites
- one `ADDIN ZOTERO_BIBL` field replacing the old reference list, which Zotero fills in on refresh
- document preferences (`ZOTERO_PREF_*` custom properties) naming the chosen citation style

## Before starting: check the ground

Confirm these first — each one, if wrong, produces a confusing failure later:

1. **Zotero is running** and its local API answers: `curl -s -o /dev/null -w '%{http_code}' http://127.0.0.1:23119/api/users/0/items?limit=1` should print `200`.
2. **The Zotero Word plugin is installed** — a `Zotero.dotm` in `%APPDATA%\Microsoft\Word\STARTUP\` (Windows) or the equivalent. Without it the fields can never be refreshed.
3. **The user edits in Microsoft Word, not WPS.** The plugin registers with Word; WPS will not run it. If the document was authored in WPS that is fine (the file format is the same) as long as they open it in Word.
4. **Read the document's current citation form** and confirm it is the supported one — numbered brackets `[1,2]` in the body plus a numbered reference list at the end. If the body uses author–date `(Smith, 2020)`, superscript numbers, or has no in-text markers at all, stop and tell the user this skill does not cover that form yet; do not improvise a parser, because a partial parse silently drops citations and that failure is invisible in the output.

Also ask which **citation style** to write into the document (e.g. `elsevier-harvard`, `american-chemical-society`, `china-national-standard-gb-t-7714-2015-numeric`). Getting this wrong changes how every citation looks, so ask rather than guess. Style IDs follow `http://www.zotero.org/styles/<name>`; `ls "$ZOTERO_DATA/styles/"` shows what is already installed locally.

## The one trap that matters most

**Never assemble the .docx from scratch, and never hand Word a minimal hand-built document.**

A minimal docx opens in Word's compatibility mode, and Zotero's refresh then *hangs Word completely* — no error dialog, no log, just a frozen application. This was reproduced both from script and by hand. The fix is to always rewrite a document that Word itself accepts: take the user's real file (or a copy) and modify its `word/document.xml` in place.

Corollary: the injected file is a rewrite of the original package. Copy every other part through unchanged — do not regenerate styles, settings, fonts or media.

## Pipeline

Three bundled scripts, run in order. They take explicit paths so nothing is hardcoded to one machine; run them from a scratch directory next to the manuscript.

### 1. Resolve references to Zotero items

```bash
python scripts/resolve.py --docx "<path to manuscript.docx>" --out-dir "<scratch dir>"
```

Parses the reference list, matches each entry against the Zotero library, and writes `refmap.json` (`{reference number: Zotero item key}`) plus `citation-map.md` for the user to audit.

Matching is two-pass because a manuscript's wording often differs slightly from the Zotero title ("in **the** water-sediment system" vs "in water sediment system"): first token-containment against library titles, then for whatever is left, ask CrossRef for the authoritative title behind the entry's DOI and match that.

**If any reference is unresolved, stop and report it to the user with its DOI.** Do not import anything into their library on your own initiative — writing to someone's reference library is their decision. If they ask you to import, the `zotero_add_item` MCP tool takes DOIs directly and can file them into a collection.

### 2. Inject the fields

```bash
python scripts/inject.py \
  --docx "<manuscript.docx>" \
  --refmap "<scratch>/refmap.json" \
  --style "http://www.zotero.org/styles/<style-id>" \
  --out "<manuscript_v2.docx>"
```

Rewrites the body citations into fields, swaps the reference list for the bibliography field, writes the document preferences, and strips the direct red character colour that flattened-citation documents usually carry as renumbering residue.

Why the body pass is not a simple find/replace: **citation tokens are frequently split across several text runs.** Red colour marks, spell-check boundaries and revision marks cut through them, so a regex over the text of individual runs misses a large fraction — in the document this skill was built from, 26 of 48 tokens were split across 2–8 runs. The script therefore joins each paragraph's runs into one string, locates tokens there, and rebuilds the run structure around them.

### 3. Verify — do not skip this

```bash
python scripts/verify.py --original "<manuscript.docx>" --output "<output.docx>" \
                         --refmap "<scratch>/refmap.json" --itemmeta "<scratch>/itemmeta.json"
```

This runs two independent checks and both matter:

- **Positional correspondence** — walking the body in order, the Nth field in the output must cite exactly what the Nth `[n]` token of the original did. Catches scrambling introduced by the run surgery.
- **Semantic correctness** — for every reference, compare the manuscript's own entry against the Zotero item's metadata (DOI first, then title/author/year). This is the check that catches a wrong mapping, which the positional check cannot see, and it frequently also exposes **errors in the manuscript itself** — DOIs that do not resolve, or that resolve to a completely different paper. Report those to the user; do not silently "fix" a reference list you were not asked to change.

Report the results honestly. If a check fails, say so and stop rather than shipping a document with mismatched citations — a wrong citation is invisible on the page and actively harmful in a submission.

### 4. Have the user refresh in Word

Structural verification proves the fields are well-formed; it cannot prove Zotero adopts them. Only Word can:

1. Open the output in Word, click **Zotero → Refresh**
2. The citations re-render and the reference list appears

Two things make this step fail for reasons that look like a broken conversion, so
check them before concluding anything is wrong:

- **Zotero's Word integration handles one document at a time.** If another Word
  window is open — especially one already holding Zotero-linked fields — Refresh
  can return without error and change nothing at all. Ask the user to close every
  other Word window first.
- **Driving Word by script is not a substitute.** Launching a separate hidden Word
  instance and calling the `ZoteroRefresh` macro can silently do nothing, because
  the add-in is registered with the user's normal instance. Prefer having the user
  do it; if you must automate, drive their existing instance rather than a new one,
  and never close a window you did not open.

Tell the user to check the refresh happened and that the style is what they
wanted. Since the fields stay live, changing style afterwards is **Zotero →
Document Preferences**, and adding citations afterwards is **Zotero → Add/Edit
Citation** — not typing `[n]` by hand. Warn them off **Unlink Citations**, which
flattens the fields back to text and undoes the whole exercise.

## Field format reference

Read `references/field-format.md` when you need the exact field syntax, the `ZOTERO_PREF` document-property layout, or are debugging why Zotero refuses to adopt the fields.

## Limits

- Numbered-bracket citations only (`[1,2]`). Author–date and superscript forms are not parsed.
- Windows-oriented: the Word-automation hints and plugin paths assume it. The docx rewriting itself is platform-neutral.
- PDFs are not attached; only metadata is used.
