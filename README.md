# Table of Authorities for LibreOffice Writer

Writer has no table of authorities. This adds one: mark the citations in a
brief, and it generates a categorised table with page numbers.

```
TABLE OF AUTHORITIES                                    Page(s)
CASES
Doe v. Acme Insurance Co., 329 So. 3d 148                     2
Smith v. Jones, 260 So. 3d 323                           passim
STATUTES
section 627.736, Florida Statutes                             4
RULES
Fla. R. Civ. P. 1.510(c)                                      3
```

Status: **early**. The marking and generation work and are tested end to end
against a real LibreOffice; the output is a plain table rather than the tab
stops with dot leaders a filed brief wants. See [Limitations](#limitations).

## Why this exists

Writer *can* make several user-defined indexes in one document — the common
claim that it cannot is wrong. What it cannot do is sort them: a user-defined
index lists entries in the order they appear and will not merge duplicates
into a page list, which is exactly what a table of authorities is. The
alphabetical index does sort, but a document may only have one, and a brief
needs cases, statutes and rules listed separately.

This extension does the sorting, grouping and page collection itself, and
writes the result into the document.

## Install

Download the `.oxt` from
[Releases](https://github.com/harkins-solutions/libreoffice-toa/releases), then
**Tools > Extension Manager > Add**. Restart Writer. An **Authorities** menu
appears.

Or build it yourself — there is nothing to compile:

```console
$ python3 build.py            # writes dist/table-of-authorities.oxt
```

## Use

1. Select a citation in your brief.
2. **Authorities > Mark Citation...** — confirm the text and pick a category
   (Cases, Statutes, Rules, Constitutional Provisions, Other).
3. Mark every place the authority appears, including short forms and `id.`
   cites. Give them the **same authority text** as the full citation and they
   group into one entry with all their pages.
4. Put the cursor where the table belongs and choose
   **Authorities > Insert Table of Authorities**.

**Authorities > Show Marked Citations** lists what is currently marked.

Page numbers are read from the document's layout when the table is generated,
so regenerate after the brief's pagination changes.

## How a mark is stored

Two things, both inside the document:

| What | Where | Carries |
|---|---|---|
| a bookmark named `toa0000`, `toa0001`… | over the cited text | the position, and therefore the page |
| JSON in the `ToaMarks` user-defined document property | document properties | the authority text and its category |

Both survive `.odt` and `.docx` round trips, and bookmarks move with the text,
so marks stay correct as the brief is edited. `tests/run_tests.py` checks this
against a real LibreOffice rather than assuming it.

Behaviour worth knowing:

- Marking the same range twice does not create a second mark. Marking it again
  with a different category updates the existing one.
- Deleting cited text from the brief drops it from the table. A bookmark
  survives the deletion of the text it covered, so the extension treats an
  empty anchor as "no longer cited" rather than listing an authority the brief
  does not contain.
- The table shows the authority text you recorded, not the current document
  text. That is what lets a short form group under its full citation; it also
  means editing the cited words does not change the entry. Re-mark to update.
- Five or more pages renders as `passim`. Consecutive pages collapse: `1-3, 5`.

## Limitations

- The table is a Writer table. A filed brief wants tab stops with dot leaders.
- Marking is manual. Detecting citations automatically is planned as a
  *suggestion* step that proposes marks for confirmation — never one that
  writes the table directly, because a missed authority in a filed table of
  authorities is worse than an empty one.
- No settings yet: `passim` at five pages and the category list are fixed.
- Not tested against Microsoft Word itself. Bookmarks and custom document
  properties are both standard OOXML, and they survive LibreOffice's own
  `.docx` round trip, but whether Word preserves them is untested.

## Tests

```console
$ tests/run.sh
```

Builds the extension, installs it with `unopkg`, and drives the same dispatch
URLs the menu uses against a headless LibreOffice — so what is tested is the
extension as installed, not the source imported directly. Set `LO_PROGRAM` if
your LibreOffice `program` directory is somewhere unusual.

## License

Apache-2.0.
