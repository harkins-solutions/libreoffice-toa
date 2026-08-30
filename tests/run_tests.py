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
        d = doc.createSearchDescriptor()
        d.SearchString = needle
        d.SearchCaseSensitive = True
        found = doc.findAll(d)
        if index >= found.getCount():
            return None
        rng = found.getByIndex(index)
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


def main():
    session = Session()
    try:
        test_duplicate_marking(session)
        test_remark_updates_category(session)
        test_distinct_occurrences_are_separate_marks(session)
        test_table_contents(session)
        test_page_ranges(session)
        test_deleting_the_text_drops_the_entry(session)
    finally:
        session.close()

    passed = sum(1 for _, ok, _, _ in RESULTS if ok)
    print(f"\n{'=' * 60}\n{passed}/{len(RESULTS)} checks passed")
    return 0 if passed == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
