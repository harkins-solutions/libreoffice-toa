"""End-to-end tests against a real LibreOffice, driving the installed add-on.

Run with LibreOffice's bundled python, after `unopkg add dist/*.oxt`:

    tests/run.sh

Every test dispatches the same URLs the Authorities menu dispatches, so what
is tested is the extension as installed, not the source imported directly.
"""
import subprocess
import sys
import time
from pathlib import Path

SOCKET = "socket,host=127.0.0.1,port=2010;urp;StarOffice.ComponentContext"
OUT = Path(__file__).resolve().parent / "out"

RESULTS = []


def check(name, got, want):
    ok = got == want
    RESULTS.append((name, ok, got, want))
    print(f"  {'PASS' if ok else 'FAIL'}  {name}"
          + ("" if ok else f"   got {got!r}, wanted {want!r}"))
    return ok


def connect(retries=40):
    import uno
    from com.sun.star.connection import NoConnectException
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(retries):
        try:
            return resolver.resolve(f"uno:{SOCKET}")
        except NoConnectException:
            time.sleep(1)
    raise SystemExit("soffice did not accept a connection")


def prop(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


class Session:
    def __init__(self):
        self.proc = subprocess.Popen(
            ["soffice", "--headless", "--norestore", "--nologo",
             f"--accept={SOCKET}"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        self.ctx = connect()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)
        self.helper = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", self.ctx)

    def new_doc(self, pages):
        import uno
        doc = self.desktop.loadComponentFromURL(
            "private:factory/swriter", "_blank", 0, ())
        text = doc.getText()
        cursor = text.createTextCursor()
        for number, lines in enumerate(pages, start=1):
            text.insertString(cursor, f"Argument {number}.\n", False)
            for line in lines:
                text.insertString(cursor, f"See {line} at bar.\n", False)
            if number < len(pages):
                text.insertControlCharacter(cursor, 0, False)
                cursor.setPropertyValue("BreakType", uno.Enum(
                    "com.sun.star.style.BreakType", "PAGE_BEFORE"))
        return doc

    def dispatch(self, doc, command, **args):
        frame = doc.getCurrentController().getFrame()
        self.helper.executeDispatch(
            frame, f"org.libreoffice.toa:{command}", "", 0,
            tuple(prop(k, v) for k, v in args.items()))

    def select(self, doc, needle, index=0):
        """Select the index-th occurrence in the body text.

        Once a table has been generated it contains a copy of every authority,
        and findAll returns those too -- not in document order, so the first
        hit may be the one inside the table. A user selects what they can see;
        a test has to say which it means.
        """
        d = doc.createSearchDescriptor()
        d.SearchString = needle
        d.SearchCaseSensitive = True
        found = doc.findAll(d)
        body = doc.getText()
        in_body = [found.getByIndex(i) for i in range(found.getCount())
                   if found.getByIndex(i).getText() == body]
        if index >= len(in_body):
            return None
        rng = in_body[index]
        doc.getCurrentController().select(rng)
        return rng

    def marks(self, doc):
        import json
        props = doc.getDocumentProperties().getUserDefinedProperties()
        try:
            return json.loads(props.getPropertyValue("ToaMarks") or "{}")
        except Exception:
            return {}

    def close(self):
        self.proc.terminate()


SMITH = "Smith v. Jones, 260 So. 3d 323"
DOE = "Doe v. Acme Insurance Co., 329 So. 3d 148"
RULE = "Fla. R. Civ. P. 1.510(c)"
STATUTE = "section 627.736, Florida Statutes"


def test_duplicate_marking(s):
    """The bug: marking the same range twice piled up bookmarks."""
    print("\ntest: marking the same range twice")
    doc = s.new_doc([[SMITH], [DOE]])
    s.select(doc, SMITH)
    s.dispatch(doc, "MarkCitation", Authority=SMITH, Category="Cases")
    first = len(s.marks(doc))
    bookmarks_first = doc.getBookmarks().getCount()

    s.select(doc, SMITH)          # same range again
    s.dispatch(doc, "MarkCitation", Authority=SMITH, Category="Cases")

    check("one mark after marking once", first, 1)
    check("still one mark after marking twice", len(s.marks(doc)), 1)
    check("still one bookmark", doc.getBookmarks().getCount(), bookmarks_first)
    doc.close(False)


def test_remark_updates_category(s):
    print("\ntest: re-marking with a different category updates it")
    doc = s.new_doc([[STATUTE]])
    s.select(doc, STATUTE)
    s.dispatch(doc, "MarkCitation", Authority=STATUTE, Category="Cases")
    s.select(doc, STATUTE)
    s.dispatch(doc, "MarkCitation", Authority=STATUTE, Category="Statutes")
    marks = s.marks(doc)
    check("one mark, not two", len(marks), 1)
    check("category updated", list(marks.values())[0]["category"], "Statutes")
    doc.close(False)


def test_distinct_occurrences_are_separate_marks(s):
    """Two occurrences of one authority are two marks, on two pages."""
    print("\ntest: the same authority cited on two pages")
    doc = s.new_doc([[SMITH], [SMITH]])
    for i in (0, 1):
        s.select(doc, SMITH, index=i)
        s.dispatch(doc, "MarkCitation", Authority=SMITH, Category="Cases")
    check("two marks", len(s.marks(doc)), 2)
    doc.close(False)


def test_table_contents(s):
    print("\ntest: the generated table")
    pages = [[SMITH], [DOE], [SMITH, RULE], [SMITH, STATUTE], [SMITH], [SMITH]]
    doc = s.new_doc(pages)
    for authority, category in ((SMITH, "Cases"), (DOE, "Cases"),
                                (RULE, "Rules"), (STATUTE, "Statutes")):
        d = doc.createSearchDescriptor()
        d.SearchString = authority
        found = doc.findAll(d)
        for i in range(found.getCount()):
            s.select(doc, authority, index=i)
            s.dispatch(doc, "MarkCitation", Authority=authority,
                       Category=category)
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")

    tables = doc.getTextTables()
    check("one table produced", tables.getCount(), 1)
    if tables.getCount():
        table = tables.getByIndex(0)
        rows = table.getRows().getCount()
        cells = {}
        for r in range(1, rows + 1):
            cells[table.getCellByName(f"A{r}").getString()] = \
                table.getCellByName(f"B{r}").getString()
        check("Smith collapses to passim (6 pages)", cells.get(SMITH), "passim")
        check("Doe shows its single page", cells.get(DOE), "2")
        check("rule page", cells.get(RULE), "3")
        check("statute page", cells.get(STATUTE), "4")
        check("categories present",
              [c for c in ("CASES", "STATUTES", "RULES") if c in cells],
              ["CASES", "STATUTES", "RULES"])
    OUT.mkdir(exist_ok=True)
    doc.storeToURL((OUT / "table.odt").as_uri(), ())
    doc.close(False)


def test_page_ranges(s):
    print("\ntest: consecutive pages collapse to a range")
    doc = s.new_doc([[DOE], [DOE], [DOE], ["filler"], [DOE]])
    d = doc.createSearchDescriptor()
    d.SearchString = DOE
    for i in range(doc.findAll(d).getCount()):
        s.select(doc, DOE, index=i)
        s.dispatch(doc, "MarkCitation", Authority=DOE, Category="Cases")
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")
    table = doc.getTextTables().getByIndex(0)
    label = None
    for r in range(1, table.getRows().getCount() + 1):
        if table.getCellByName(f"A{r}").getString() == DOE:
            label = table.getCellByName(f"B{r}").getString()
    check("pages 1,2,3 and 5 render as '1-3, 5'", label, "1-3, 5")
    doc.close(False)


def test_deleting_the_text_drops_the_entry(s):
    print("\ntest: deleting cited text removes it from the table")
    doc = s.new_doc([[SMITH], [DOE]])
    for authority in (SMITH, DOE):
        s.select(doc, authority)
        s.dispatch(doc, "MarkCitation", Authority=authority, Category="Cases")
    rng = s.select(doc, SMITH)
    rng.setString("")                       # the author deletes the citation
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")
    table = doc.getTextTables().getByIndex(0)
    listed = [table.getCellByName(f"A{r}").getString()
              for r in range(1, table.getRows().getCount() + 1)]
    check("deleted authority is gone", SMITH in listed, False)
    check("the other authority remains", DOE in listed, True)
    doc.close(False)


def test_regenerating_updates_one_table(s):
    """Generating twice used to leave two tables disagreeing with each other."""
    print("\ntest: generating the table twice")
    doc = s.new_doc([[SMITH], [DOE]])
    s.select(doc, SMITH)
    s.dispatch(doc, "MarkCitation", Authority=SMITH, Category="Cases")
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")
    check("one table after generating", doc.getTextTables().getCount(), 1)

    # Mark another authority and regenerate, as anyone drafting would.
    s.select(doc, DOE)
    s.dispatch(doc, "MarkCitation", Authority=DOE, Category="Cases")
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")
    check("still one table after regenerating", doc.getTextTables().getCount(), 1)

    table = doc.getTextTables().getByIndex(0)
    listed = [table.getCellByName(f"A{r}").getString()
              for r in range(1, table.getRows().getCount() + 1)]
    check("the new authority is in it", DOE in listed, True)
    check("the first authority is still in it", SMITH in listed, True)
    doc.close(False)


def test_the_table_shrinks_as_well_as_grows(s):
    print("\ntest: the table shrinks when an authority goes away")
    doc = s.new_doc([[SMITH], [DOE]])
    for authority in (SMITH, DOE):
        s.select(doc, authority)
        s.dispatch(doc, "MarkCitation", Authority=authority, Category="Cases")
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")
    before = doc.getTextTables().getByIndex(0).getRows().getCount()

    s.select(doc, DOE)
    s.dispatch(doc, "UnmarkCitation")
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")
    table = doc.getTextTables().getByIndex(0)
    listed = [table.getCellByName(f"A{r}").getString()
              for r in range(1, table.getRows().getCount() + 1)]
    check("a row was removed", table.getRows().getCount() < before, True)
    check("the unmarked authority is gone", DOE in listed, False)
    check("the other one remains", SMITH in listed, True)
    doc.close(False)


def test_unmark_removes_the_mark_and_the_bookmark(s):
    print("\ntest: unmarking a citation")
    doc = s.new_doc([[SMITH], [DOE]])
    for authority in (SMITH, DOE):
        s.select(doc, authority)
        s.dispatch(doc, "MarkCitation", Authority=authority, Category="Cases")
    check("two marks to start", len(s.marks(doc)), 2)
    bookmarks_before = doc.getBookmarks().getCount()

    s.select(doc, SMITH)
    s.dispatch(doc, "UnmarkCitation")
    check("one mark left", len(s.marks(doc)), 1)
    check("the bookmark went too", doc.getBookmarks().getCount(),
          bookmarks_before - 1)
    check("the right one survived",
          [m["authority"] for m in s.marks(doc).values()], [DOE])
    # The text itself must not be touched.
    d = doc.createSearchDescriptor()
    d.SearchString = SMITH
    check("the citation text is still in the document",
          doc.findAll(d).getCount(), 1)
    doc.close(False)


def test_unmark_with_the_cursor_inside_the_citation(s):
    """A user puts the cursor in the citation; they do not select it exactly."""
    print("\ntest: unmarking with only a cursor inside the text")
    doc = s.new_doc([[SMITH]])
    s.select(doc, SMITH)
    s.dispatch(doc, "MarkCitation", Authority=SMITH, Category="Cases")

    rng = s.select(doc, SMITH)
    view = doc.getCurrentController().getViewCursor()
    view.gotoRange(rng, False)
    view.goRight(4, False)          # inside the marked text, nothing selected
    s.dispatch(doc, "UnmarkCitation")
    check("the mark is gone", len(s.marks(doc)), 0)
    doc.close(False)


def test_unmark_where_nothing_is_marked(s):
    print("\ntest: unmarking where there is no mark")
    doc = s.new_doc([[SMITH], [DOE]])
    s.select(doc, SMITH)
    s.dispatch(doc, "MarkCitation", Authority=SMITH, Category="Cases")
    s.select(doc, DOE)              # never marked
    s.dispatch(doc, "UnmarkCitation")
    check("the other mark is untouched", len(s.marks(doc)), 1)
    doc.close(False)


def test_unmarking_from_inside_the_generated_table_does_nothing(s):
    """Selecting an entry in the table is not the same as the citation."""
    print("\ntest: unmark attempted from inside the generated table")
    doc = s.new_doc([[SMITH]])
    s.select(doc, SMITH)
    s.dispatch(doc, "MarkCitation", Authority=SMITH, Category="Cases")
    doc.getCurrentController().getViewCursor().gotoEnd(False)
    s.dispatch(doc, "InsertTable")

    d = doc.createSearchDescriptor()
    d.SearchString = SMITH
    found = doc.findAll(d)
    body = doc.getText()
    in_table = [found.getByIndex(i) for i in range(found.getCount())
                if found.getByIndex(i).getText() != body]
    check("the table holds a copy of the authority", len(in_table), 1)
    if in_table:
        doc.getCurrentController().select(in_table[0])
        s.dispatch(doc, "UnmarkCitation")
    check("the mark is untouched", len(s.marks(doc)), 1)
    doc.close(False)


def main():
    session = Session()
    try:
        test_duplicate_marking(session)
        test_remark_updates_category(session)
        test_distinct_occurrences_are_separate_marks(session)
        test_table_contents(session)
        test_page_ranges(session)
        test_deleting_the_text_drops_the_entry(session)
        test_regenerating_updates_one_table(session)
        test_the_table_shrinks_as_well_as_grows(session)
        test_unmark_removes_the_mark_and_the_bookmark(session)
        test_unmark_with_the_cursor_inside_the_citation(session)
        test_unmark_where_nothing_is_marked(session)
        test_unmarking_from_inside_the_generated_table_does_nothing(session)
    finally:
        session.close()

    passed = sum(1 for _, ok, _, _ in RESULTS if ok)
    print(f"\n{'=' * 60}\n{passed}/{len(RESULTS)} checks passed")
    return 0 if passed == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
