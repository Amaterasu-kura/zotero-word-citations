# Zotero Word field format

Everything here was read out of real Zotero-managed documents and Zotero's own
`xpcom/integration.js`, and confirmed by round-tripping a 70-reference manuscript.

## Citation field

The whole instruction lives in a single `w:instrText` run, wrapped in `w:fldChar`
begin / separate / end runs. Note the **leading and trailing space** inside the
instruction — Zotero's parser expects them.

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {...json...} </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>(Galand et al., 2018)</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

The JSON:

```json
{
  "citationID": "<8 random alphanumerics>",
  "properties": {
    "unsorted": false,
    "formattedCitation": "(Galand et al., 2018)",
    "plainCitation": "(Galand et al., 2018)",
    "noteIndex": 0
  },
  "citationItems": [
    {
      "id": 1695,
      "uris": ["http://zotero.org/users/9287589/items/U8FKQ3QC"],
      "itemData": { ...full CSL-JSON for the item... }
    }
  ],
  "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
}
```

Notes that matter:

- **`uris` is what Zotero resolves on refresh**; `itemData` is the fallback copy.
  Get the URI from Zotero itself — `GET /api/users/0/items/<KEY>?format=csljson`
  returns an `id` field that *is* the correct URI. Do not construct it by hand;
  the user id in it is the account's, not something you can guess.
- **`id` is the numeric `items.itemID`**, which the local API does not expose. Read
  it from a WAL-inclusive snapshot copy of `zotero.sqlite` (Zotero holds an
  exclusive lock on the live file, so copy `zotero.sqlite`, `-wal` and `-shm`
  together and open the copy read-only).
- Multiple works in one citation go into a single field's `citationItems` array —
  Zotero then renders "(A, 2018; B, 2020)". Do not emit several adjacent fields.
- The `separate`…`end` region holds the rendered text. Zotero overwrites it on
  refresh, so an approximate string is fine and better than a placeholder: an
  author–date document still reads correctly before the user refreshes.

## Bibliography field

```xml
<w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY </w:instrText>
```

Same begin/separate/end wrapping. Replace the whole plain-text reference list —
every entry paragraph — with a single paragraph holding this field.

## Document preferences: `ZOTERO_PREF_*`

**Zotero does not store the citation style in the fields.** It stores document
preferences in custom document properties, and without them Zotero will not treat
the file as one of its own. Office caps a custom-property string at 255 characters,
so the payload is split across `ZOTERO_PREF_1`, `ZOTERO_PREF_2`, …:

```xml
<data data-version="3" zotero-version="6.0.7">
  <session id="XXXXXXXX"/>
  <style id="http://www.zotero.org/styles/elsevier-harvard"
         hasBibliography="1" bibliographyStyleHasBeenSet="1"/>
  <prefs>
    <pref name="fieldType" value="Field"/>
    <pref name="automaticJournalAbbreviations" value="true"/>
  </prefs>
</data>
```

Three package-level pieces must all be present, or Word silently drops the
properties:

1. the `docProps/custom.xml` part itself
2. an `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>` in `[Content_Types].xml`
3. a relationship of type `.../officeDocument/2006/relationships/custom-properties` targeting `docProps/custom.xml` in `_rels/.rels`

`data-version` must be **3** — that is `DATA_VERSION` in Zotero's `integration.js`.
A higher value raises "newer document version"; 4 is tolerated as a transitional
JSON form.

## Why Zotero hangs Word, and what to do

A hand-built minimal .docx opens in Word's compatibility mode and makes the Zotero
refresh freeze Word entirely — no dialog, no log. Always start from a document Word
itself has written. If a refresh hangs and you need to recover, kill the WINWORD
process (check its window title first: it should hold only your test file) and
delete any stale `~$<name>.docx` lock file next to the document.

## Driving the refresh from script

For end-to-end testing without a human:

```python
import win32com.client as wc
app = wc.Dispatch("Word.Application")
doc = app.Documents.Open(path)
app.Run("ZoteroRefresh")          # <-- this name
doc.SaveAs2(out, FileFormat=16)
```

`ZoteroRefresh` is the macro; the ribbon callback `ZoteroRibbon.ZoteroRibbonRefresh`
is *not* callable from COM because it requires an `IRibbonControl` argument.
Enumerating the VBA project is blocked unless the user has enabled "Trust access to
the VBA project object model" — do not ask them to change a security setting just to
run a test.

## Failure modes seen in practice

| Symptom | Cause |
|---|---|
| Word freezes on Refresh, no error | document was hand-built rather than Word-authored |
| Zotero ignores the fields entirely | missing `ZOTERO_PREF_*` custom properties |
| `ZoteroRefresh` returns cleanly but nothing changes | Zotero's Word integration is already holding another open document; it handles one at a time. Close every other Word window and retry |
| Fields become "unlinked" after refresh | `uris` wrong — usually a hand-built URI or a foreign user id |
| Citation renders but points at the wrong paper | reference→item mapping error; caught by `verify.py`'s semantic check |
| Bibliography shows stale text | bibliography field not inserted, or old list paragraphs left in place |
| Body text reads as the literal `&lt;16 µm`, `&gt;400,000` | the run rewrite re-escaped text that was already XML-escaped. Only hits runs that carry an entity *and* part of a citation token, so it is easy to miss — `verify.py` checks for it |
| Pre-existing custom properties vanished | `docProps/custom.xml` was replaced instead of merged. Merge; authoring tools put workflow metadata there and losing it is silent |

## Rewriting run text without corrupting it

`<w:t>` content is XML-escaped on the way in (`<` is stored as `&lt;`). When you
split a run and re-emit its text, insert the string exactly as you read it — do
**not** escape it again, or `&lt;16 µm` becomes `&amp;lt;16 µm` and Word prints
the entity literally. Escape only text *you* generated.

The trap is narrow enough to slip through review: it fires solely on runs that
both contain an entity and are cut by a citation token. Check for it
mechanically rather than by eye, and check only visible text — entities inside
`<w:instrText>` are the field's JSON payload (Zotero stores some item metadata
pre-escaped) and are never displayed.

