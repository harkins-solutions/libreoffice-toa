# Contributing

## Never send a real brief

Not in an issue, not in a pull request, not in a test fixture. Invented
citations reproduce every bug this extension can have, because it sorts and
groups the text it is given and never looks anything up.

## Build it

There is nothing to compile. An `.oxt` is a zip.

```console
$ python3 build.py                 # writes dist/table-of-authorities.oxt
$ unopkg add -f dist/table-of-authorities.oxt
```

Restart Writer and the **Authorities** menu appears. `unopkg add -f` replaces
an already-installed copy, so that one line is the whole edit-and-try loop.

## Test it

```console
$ tests/run.sh                     # builds, installs, runs against real Writer
```

The tests drive a headless LibreOffice through the same dispatch URLs the menu
uses, so they exercise the extension as installed rather than the source
imported directly. There are no mocks: a component that registers but fails to
dispatch, or a page number that is wrong, only shows up against the real
application.

`tests/run.sh` finds LibreOffice in `$LO_PROGRAM`, then a system install, then
a local extracted copy. Set `LO_PROGRAM` if yours is somewhere else.

On Windows with Word installed, `tests/word_roundtrip.sh` additionally opens a
marked `.docx` in Word through COM automation and checks the marks survive.
It skips itself elsewhere.

## What a change should come with

A test that fails before it and passes after. The suite is small and fast
because everything in it is there for a reason — each check corresponds to
something that was once wrong.

Two bugs found this way are worth knowing about, because they are the shape of
what goes wrong here:

- Marking the same range twice piled up bookmarks. The table still looked
  right, because entries group by authority, so nothing surfaced until later
  edits disagreed about what was marked.
- Deleting a citation left it in the table. A bookmark outlives the text it
  covered, collapsing to a point, and still reports a page.

Both were invisible in the output and obvious in a test.

## Style

Python that reads like the file it sits in. Comments explain why something is
the way it is, not what the line does — particularly where a UNO API behaves
in a way you would not predict. `_disjoint()` carries the comparison
convention it depends on, because that convention was measured rather than
read off the documentation, and the next person should not have to measure it
again.

## Scope

This generates a table of authorities. It is not a citation checker, a
formatter, or a brief-drafting tool. If a change would make it any of those,
it probably belongs in a different project.
