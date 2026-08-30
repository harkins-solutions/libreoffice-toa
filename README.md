# Table of Authorities for LibreOffice Writer

Writer has no table of authorities. This adds one: mark the citations in a
brief, and it generates the table, categorised, alphabetised, with page
numbers.

```
TABLE OF AUTHORITIES                                    Page(s)
CASES
Doe v. Acme Insurance Co., 329 So. 3d 148                     2
Smith v. Jones, 260 So. 3d 323                           passim
STATUTES
section 627.736, Florida Statutes                             4
RULES
Fla. R. Civ. P. 1.510(c)                                    1-3, 5
```

Python, no compilation, no network access, nothing stored outside your
document. LibreOffice 4.1 or newer. Apache-2.0.

**Status: early.** Marking and generation work and are tested end to end
against a real LibreOffice and against Microsoft Word. The output is a plain
table rather than the tab stops with dot leaders a filed brief wants — see
[Limitations](#limitations).

---

## Contents

[Why](#why-this-exists) ·
[Install](#install) ·
[Use](#use) ·
[What the table looks like](#what-the-table-looks-like) ·
[How marks are stored](#how-marks-are-stored) ·
[Word compatibility](#word-compatibility) ·
[Limitations](#limitations) ·
[Tests](#tests) ·
[Contributing](CONTRIBUTING.md)

## Why this exists

The common explanation — that Writer cannot make more than one index — is
wrong. Writer will make as many *user-defined* indexes in one document as you
like. Two real limits stop that being a table of authorities:

1. A user-defined index lists entries **in the order they appear** and will not
   sort them or merge duplicates into a page list.
2. The *alphabetical* index does sort, but a document may contain **only one**,
   and a brief needs cases, statutes and rules listed separately.

So the sorting, grouping, page collection and `passim` handling have to happen
somewhere else. That is what this does; the result is written into the document
as ordinary content.

## Install

Download `table-of-authorities.oxt` from
[Releases](https://github.com/harkins-solutions/libreoffice-toa/releases), then
in Writer: **Tools > Extension Manager > Add**, select the file, and restart.
An **Authorities** menu appears.

To build it yourself — there is nothing to compile:

```console
$ python3 build.py
$ unopkg add -f dist/table-of-authorities.oxt
```

## Use

| Step | Command |
|---|---|
| 1. Select a citation in the text | — |
| 2. Mark it, choosing a category | **Authorities > Mark Citation...** |
| — undo a mistaken mark: put the cursor in it | **Authorities > Unmark Citation** |
| 3. Repeat for every appearance, including short forms and `id.` cites | same |
| 4. Put the cursor where the table belongs | — |
| 5. Generate, or bring an existing table up to date | **Authorities > Insert or Update Table of Authorities** |

**Authorities > Show Marked Citations** lists what is currently marked.

Short forms group under their full citation when you give them the **same
authority text**. Marking `Smith, 260 So. 3d at 325` with the authority
`Smith v. Jones, 260 So. 3d 323` adds that page to Smith's entry rather than
creating a second one. The dialog pre-fills the selected text, so this is a
matter of correcting it before pressing Mark.

Page numbers are read from the document's layout at the moment you generate,
so regenerate after the pagination changes. Regenerating rewrites the existing
table rather than adding a second one, and keeps any formatting you applied to
it: rows are added or removed as the list grows and shrinks.

## What the table looks like

Five categories, in this order, and only the ones you have used appear:

`Cases` · `Statutes` · `Rules` · `Constitutional Provisions` · `Other`

Entries are alphabetised within each category, ignoring case. Pages are
collected across every mark for that authority:

| Pages marked | Shown as |
|---|---|
| one page | `4` |
| separate pages | `2, 7` |
| consecutive pages | `1-3` |
| a mix | `1-3, 5` |
| five or more pages | `passim` |

## How marks are stored

Two things, both inside the document itself:

| What | Where | Carries |
|---|---|---|
| a bookmark named `toa0000`, `toa0001`… | over the cited text | the position, and so the page |
| JSON in the `ToaMarks` custom document property | document properties | the authority text and category |

Behaviour that follows from this, and is tested:

- **Marks move with the text.** Insert a page above a citation and its entry
  reports the new page. Bookmarks are maintained by Writer, so this is true
  through any edit.
- **Marking the same range twice does nothing.** Marking it again with a
  different category updates the existing mark rather than adding one.
- **Deleting cited text drops it from the table.** A bookmark outlives the text
  it covered, collapsing to a point; an empty anchor is treated as "no longer
  cited", because listing an authority the brief does not contain would
  misstate the record.
- **The table shows the authority text you recorded**, not the current document
  text. That is what lets a short form group under its full citation. It also
  means editing the cited words does not change the entry — re-mark to update.
- **Unmark works on the citation, not on the table.** The generated table
  contains a copy of every authority; selecting one there and unmarking does
  nothing, because the mark lives on the citation in your text.

Anyone with the file can read the marks; see [SECURITY.md](SECURITY.md) if the
document is going to opposing counsel.

## Word compatibility

Marks survive Microsoft Word. `tests/word_roundtrip.py` opens a marked `.docx`
in Word through COM automation, has Word save it, and reopens the result:

```
PASS  Word itself reads the ToaMarks property
PASS  custom property survived Word
PASS  mark count unchanged
PASS  bookmark count unchanged
PASS  page numbers unchanged

Word's own view: 3 bookmark(s), properties: ToaMarks
```

Both constructs are standard OOXML, so a brief can move between Writer and
Word without losing what has been marked. Word does not *generate* this table
— it has its own TA fields, which this does not read or write.

## Limitations

- The table is a Writer table. A filed brief wants tab stops with dot leaders.
- Marking is manual. Automatic detection is planned as a step that **proposes**
  marks for confirmation, never one that writes the table directly: a missing
  authority in a filed table of authorities is worse than no table at all.
- No settings yet. `passim` at five pages and the category list are fixed.
- No way to jump from a table entry back to the citation in the text.
- No keyboard shortcut for marking, which is the action you repeat most.
- Word's own TA fields are neither read nor written, so a table marked up in
  Word does not carry over.

## Tests

```console
$ tests/run.sh                 # 15 checks against a headless LibreOffice
$ tests/word_roundtrip.sh      # WSL + Word only; skips itself elsewhere
```

The tests drive the installed extension through the same dispatch URLs the menu
uses, so they exercise what a user actually gets. CI installs LibreOffice and
runs them on every push.

## License

Apache-2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).

Not affiliated with or endorsed by The Document Foundation.
