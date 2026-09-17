# zotero-word-citations

> 中文版: [README.md](README.md)

Turn a Word manuscript whose citations are typed plain text into one with **live, updatable Zotero citation fields**, and regenerate its reference list from Zotero.

Concretely: every in-text `[1,2]` becomes a Zotero `ZOTERO_ITEM` field, the reference list at the end becomes a `ZOTERO_BIBL` field, and the citation style is stored in the document's `ZOTERO_PREF_*` preferences. After the conversion, one click on **Zotero → Refresh** in Word renumbers every citation, and switching style is a single **Document Preferences** dialog — two things a plain-text reference list can never do.

This repository is an **agent skill** (`SKILL.md` carries the triggers and the operating rules) plus three Python scripts that perform *resolve → inject → verify*.

---

## What you get

- One `ADDIN ZOTERO_ITEM CSL_CITATION` field per in-text citation, each carrying the Zotero item URIs it cites
- One `ADDIN ZOTERO_BIBL` field replacing the old reference list, which Zotero fills in on Refresh
- Document preferences (`ZOTERO_PREF_*` custom properties) naming the chosen citation style
- **The original file is never overwritten** — a new document is written

## Prerequisites

| Requirement | Check / notes |
|---|---|
| Zotero | Running, with its local API enabled: **Settings → Advanced → "Allow other applications on this computer to communicate with Zotero"**. Verify with `curl -s -o /dev/null -w '%{http_code}' "http://127.0.0.1:23119/api/users/0/items?limit=1"` → must print `200` |
| Zotero Word plugin | `Zotero.dotm` installed (Windows: `%APPDATA%\Microsoft\Word\STARTUP\`). Without it the injected fields can never be refreshed |
| Microsoft Word | Editing and refreshing must happen in **Word**; **WPS does not load the plugin** (a manuscript authored in WPS is fine as long as it is opened in Word) |
| **zotero-mcp** | Lets the agent search and write to your Zotero library (e.g. import by DOI via `zotero_add_item`). See below |
| Python | 3.9+, **standard library only — no pip installs required** |
| Input document | Numbered bracket citations `[1,2]` in the body plus a numbered reference list at the end. Author–date `(Smith, 2020)`, superscript numbers, or documents with no in-text markers are **not supported yet** |

### Install zotero-mcp (recommended)

[zotero-mcp](https://github.com/54yyyu/zotero-mcp) connects a Zotero library to AI assistants over the Model Context Protocol. This skill relies on it when a reference is missing from the library and needs to be imported by DOI:

```bash
uv tool install zotero-mcp-server     # or: pip install zotero-mcp-server
zotero-mcp setup                      # auto-configures Claude Desktop and similar clients
```

Then enable **Settings → Advanced → Allow other applications…** in Zotero 7+. For writes, run `zotero-mcp authorize-local` once on Zotero 10+ (choose *Always Allow*), or set `ZOTERO_API_KEY` and `ZOTERO_LIBRARY_ID` to write through the web API on older versions.

> The MCP server and the "agent skill" route share one configuration; see the zotero-mcp repository for details.

## Installing this skill

Drop the whole directory into your agent's skills directory (for Claude Code that is `~/.claude/skills/zotero-word-citations/`; other clients use their own location).
The `SKILL.md` frontmatter already declares the `name` and trigger phrases ("把文献插进 Word", "引用无法更新", "参考文献表是纯文本", "换引用样式", …), so the agent loads it automatically for matching requests.

## Quick start

Run the three scripts in order. Every path is passed explicitly — nothing is hardcoded to one machine. A scratch directory next to the manuscript is a good place for the intermediate files.

```bash
# 0) Confirm the document form ([1,2] in body + numbered reference list) and pick a style ID
#    List locally installed styles:  ls "$ZOTERO_DATA/styles/"

# 1) Resolve the reference list to Zotero items
python scripts/resolve.py --docx "manuscript.docx" --out-dir "work"

# 2) Inject the fields ([n] → Zotero citation fields, bibliography field, style preference)
python scripts/inject.py \
  --docx "manuscript.docx" \
  --refmap "work/refmap.json" \
  --itemmeta "work/itemmeta.json" \
  --style "http://www.zotero.org/styles/elsevier-harvard" \
  --out "manuscript_v2.docx"

# 3) Verify (positional + semantic checks — do not skip)
python scripts/verify.py \
  --original "manuscript.docx" \
  --output "manuscript_v2.docx" \
  --refmap "work/refmap.json" \
  --itemmeta "work/itemmeta.json"
```

**4) The last step can only be done by Word**: open `manuscript_v2.docx` → click **Zotero → Refresh**. The citations re-render and the reference list is built.

> If Refresh changes nothing, **close every other Word window first** — Zotero's Word integration handles one document at a time, and a second window holding Zotero fields makes Refresh return silently without doing anything. Do not substitute a hidden scripted Word instance for the user's own refresh either.

## Scripts

### `scripts/resolve.py` — reference list → Zotero items

| Option | Meaning |
|---|---|
| `--docx` | The manuscript to process |
| `--out-dir` | Directory for intermediate artefacts |
| `--threshold` | Token-containment threshold for a title match (default `0.85`) |

Matching is two-pass: first token containment against library titles, then — for whatever is left — the entry's DOI is resolved through CrossRef to get the authoritative title and that is matched instead. (Manuscript wording often differs slightly from the Zotero title: "in **the** water-sediment system" vs "in water sediment system".)

Artefacts written to `--out-dir`:

| File | Contents |
|---|---|
| `refmap.json` | `{reference number: Zotero item key}` |
| `itemmeta.json` | `{item key: {itemID, csl}}`, consumed by `inject.py` |
| `citation-map.md` | Human-readable audit table |
| `unresolved.md` | Entries that could not be matched (with DOIs); empty when everything resolved |

> **When anything is unresolved, stop and report the DOIs to the user.** This skill never imports into someone's library on its own initiative — that is the user's decision; if they agree, use zotero-mcp's `zotero_add_item` with the DOI.

### `scripts/inject.py` — inject the Zotero fields

| Option | Meaning |
|---|---|
| `--docx` | Source manuscript |
| `--refmap` / `--itemmeta` | Artefacts from step 1 |
| `--style` | CSL style URL |
| `--out` | Output file (**must differ from the input**; the script refuses to overwrite) |
| `--keep-red` | Keep direct red character colour (stripped by default, see below) |

Besides the body pass it replaces the plain reference list with a `ZOTERO_BIBL` field, writes the document preferences, and strips the direct red character colour that flattened-citation documents usually carry as renumbering residue.

The body pass is not a find/replace: citation tokens are routinely split across several `<w:t>` runs (red marks, spell-check boundaries and revision marks cut through them). The script joins each paragraph's runs, locates tokens in the joined string, and rebuilds the run structure around them.

### `scripts/verify.py` — verification (never skip)

Two independent checks, because they fail independently:

- **A. Positional correspondence** — walking the body in order, the Nth field in the output must cite exactly what the Nth `[n]` token of the original mapped to. Catches scrambling introduced by the run surgery.
- **B. Semantic correctness** — for every reference, compare the manuscript's own entry against the Zotero item's metadata (DOI first, then the manuscript DOI resolved via CrossRef against the Zotero title, then plain title containment). Catches a wrong mapping that A cannot see, and frequently also exposes **errors in the manuscript itself** — DOIs that do not resolve, or that point at a completely different paper.

A failed check exits non-zero. Problems that belong to the manuscript (bad DOIs) are reported as `[NOTE]` and do not fail the run — report them to the user; never silently "fix" a reference list you were not asked to change.

## Citation styles (CSL style IDs)

Styles are passed as IDs of the form `http://www.zotero.org/styles/<name>`:

| Use case | Style ID |
|---|---|
| Elsevier journals (Harvard) | `elsevier-harvard` |
| American Chemical Society | `american-chemical-society` |
| China national standard GB/T 7714 (numeric) | `china-national-standard-gb-t-7714-2015-numeric` |
| Vancouver | `vancouver` |

The style changes how every citation renders, so **ask the user rather than guessing**. It can be changed afterwards in Word via **Zotero → Document Preferences**.

## Limits

- Numbered bracket citations only (`[1,2]`); author–date and superscript forms are not parsed
- Windows-oriented: the Word-automation hints and plugin paths assume it (the docx rewriting itself is platform-neutral)
- PDFs are not attached; only metadata is used

## Troubleshooting

| Symptom | Cause |
|---|---|
| Word freezes on Refresh with no error | The document was hand-built rather than Word-authored, putting Word into compatibility mode; always rewrite a real document |
| Zotero ignores the fields entirely | Missing `ZOTERO_PREF_*` custom properties (`docProps/custom.xml` absent, or not registered in `[Content_Types].xml` / `_rels/.rels`) |
| Refresh returns cleanly but nothing changes | Zotero's Word integration is already holding another open document; close every other Word window and retry |
| Fields become "unlinked" text after refresh | Wrong `uris` — a hand-built URI or a foreign user id; always take the URI from Zotero |
| Citation renders but points at the wrong paper | Reference→item mapping error; caught by `verify.py`'s semantic check |
| Bibliography still shows the old text | The bibliography field was not inserted, or old list paragraphs were left in place |
| Body text reads `&lt;16 µm`, `&gt;400,000` | A run rewrite re-escaped text that was already XML-escaped; only hits runs carrying both an entity and part of a citation token — `verify.py` checks for it |
| Pre-existing custom properties vanished | `docProps/custom.xml` was replaced instead of merged (the script preserves existing properties) |

More field-format detail (field syntax, `ZOTERO_PREF` layout, debugging) is in [`references/field-format.md`](references/field-format.md).

## Repository layout

```
zotero-word-citations/
├── SKILL.md                  # agent skill definition (triggers + operating rules)
├── README.md                 # Chinese
├── README.en.md              # English (this file)
├── LICENSE
├── scripts/
│   ├── resolve.py            # 1) reference list → Zotero items
│   ├── inject.py             # 2) inject ZOTERO_ITEM / ZOTERO_BIBL fields
│   ├── verify.py             # 3) positional + semantic verification
│   └── zcommon.py            # shared helpers (Zotero API / docx parsing, stdlib only)
├── references/
│   └── field-format.md       # Zotero Word field format and pitfalls
└── evals/
    └── evals.json            # evaluation cases
```

## How it works (key points)

Zotero's field format is well defined, but its traps cost real debugging time; this skill mostly carries those forward:

- **Never build a docx from scratch.** A minimal hand-built document opens in Word's compatibility mode, and Zotero's refresh then **freezes Word with no error** (reproduced both from script and by hand). Always rewrite `word/document.xml` inside the user's real file, copying every other part through unchanged.
- The **leading and trailing space** inside the field instruction is required; `uris` is what Zotero resolves on refresh and can only be obtained from Zotero.
- The numeric `itemID` is not exposed by the local API — read it from a WAL-inclusive snapshot copy of `zotero.sqlite` (Zotero holds an exclusive lock on the live file).
- `data-version` must be **3**; Office caps a custom-property value at 255 characters, hence the `ZOTERO_PREF_1, _2, …` chunking.

## Acknowledgements

- [zotero-mcp](https://github.com/54yyyu/zotero-mcp) by [54yyyu](https://github.com/54yyyu) — the MCP server that lets an agent search and write to Zotero; this skill depends on it for library lookup and import
- [Zotero](https://www.zotero.org/) and its Word integration (`xpcom/integration.js`) — the de-facto specification for the field format

## License

MIT License — see [LICENSE](LICENSE).
