"""Does Microsoft Word preserve the marks?

The extension stores a mark as a bookmark plus a custom document property.
Both are standard OOXML, and both survive LibreOffice's own .docx round trip
-- but LibreOffice reading its own output proves nothing about Word.

This opens the file in real Word through COM automation, has Word save it,
and then reopens the result in LibreOffice to see what came back.

Requires Windows with Word installed, reached from WSL through powershell.exe.
Skipped everywhere else. Run with LibreOffice's bundled python:

    tests/word_roundtrip.sh
"""
import json
import subprocess
import sys
import time
from pathlib import Path

SOCKET = "socket,host=127.0.0.1,port=2011;urp;StarOffice.ComponentContext"
SMITH = "Smith v. Jones, 260 So. 3d 323"
DOE = "Doe v. Acme Insurance Co., 329 So. 3d 148"
RULE = "Fla. R. Civ. P. 1.510(c)"
PLACEMENTS = [(SMITH, "Cases", 1), (DOE, "Cases", 2), (RULE, "Rules", 3)]


def windows_temp():
    out = subprocess.run(["powershell.exe", "-NoProfile", "-Command", "$env:TEMP"],
                         capture_output=True, text=True, timeout=60)
    win = out.stdout.strip()
    if not win:
        raise SystemExit("no Windows TEMP; is this WSL?")
    linux = subprocess.run(["wslpath", "-u", win], capture_output=True,
                           text=True).stdout.strip()
    return win, Path(linux)


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
    raise SystemExit("no connection to soffice")


def prop(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def read_marks(doc):
    props = doc.getDocumentProperties().getUserDefinedProperties()
    try:
        return json.loads(props.getPropertyValue("ToaMarks") or "{}")
    except Exception:
        return {}


def page_of(doc, rng):
    cursor = doc.getCurrentController().getViewCursor()
    cursor.gotoRange(rng, False)
    return cursor.getPage()


def report(doc, label):
    marks = read_marks(doc)
    bookmarks = doc.getBookmarks()
    toa_bookmarks = [n for n in bookmarks.getElementNames() if n.startswith("toa")]
    print(f"\n[{label}]")
    print(f"  ToaMarks property : {'present' if marks else 'LOST'}"
          f" ({len(marks)} mark(s))")
    print(f"  toa bookmarks     : {len(toa_bookmarks)}")
    pages = {}
    for name, meta in sorted(marks.items()):
        if not bookmarks.hasByName(name):
            print(f"  {name}: bookmark LOST")
            continue
        anchor = bookmarks.getByName(name).getAnchor()
        page = page_of(doc, anchor)
        pages[meta["authority"]] = page
        print(f"  {name}: p{page}  {meta['category']:6}  "
              f"{meta['authority'][:38]}  anchor={anchor.getString()[:28]!r}")
    return marks, len(toa_bookmarks), pages


def main():
    win_temp, linux_temp = windows_temp()
    docx_linux = linux_temp / "toa_word_test.docx"
    docx_win = f"{win_temp}\\toa_word_test.docx"

    proc = subprocess.Popen(
        ["soffice", "--headless", "--norestore", "--nologo", f"--accept={SOCKET}"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        import uno
        ctx = connect()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        helper = smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx)

        doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        text = doc.getText()
        cursor = text.createTextCursor()
        for authority, _, page in PLACEMENTS:
            text.insertString(cursor, f"Argument {page}.\n", False)
            text.insertString(cursor, f"See {authority} at bar.\n", False)
            if page < len(PLACEMENTS):
                text.insertControlCharacter(cursor, 0, False)
                cursor.setPropertyValue("BreakType", uno.Enum(
                    "com.sun.star.style.BreakType", "PAGE_BEFORE"))

        frame = doc.getCurrentController().getFrame()
        for authority, category, _ in PLACEMENTS:
            d = doc.createSearchDescriptor()
            d.SearchString = authority
            found = doc.findAll(d)
            doc.getCurrentController().select(found.getByIndex(0))
            helper.executeDispatch(
                frame, "org.libreoffice.toa:MarkCitation", "", 0,
                (prop("Authority", authority), prop("Category", category)))

        before = report(doc, "1. marked in LibreOffice")
        doc.storeToURL(docx_linux.as_uri(), (prop("FilterName", "MS Word 2007 XML"),))
        doc.close(False)
        print(f"\n  wrote {docx_linux.name} ({docx_linux.stat().st_size} bytes)")

        # --- hand it to real Word ---------------------------------------
        # CustomDocumentProperties is a COM collection PowerShell will not
        # enumerate with foreach -- it returns nothing and looks like Word
        # dropped the property. Late binding through InvokeMember reads it.
        script = (
            "$b = 'System.Reflection.BindingFlags' -as [type]; "
            "$w = New-Object -ComObject Word.Application; "
            "$w.Visible = $false; $w.DisplayAlerts = 0; "
            f"$d = $w.Documents.Open('{docx_win}'); "
            "Write-Output ('WORD_BOOKMARKS=' + $d.Bookmarks.Count); "
            "$cp = $d.CustomDocumentProperties; "
            "$n = [System.__ComObject].InvokeMember('Count', $b::GetProperty,"
            " $null, $cp, $null); "
            "$names = @(); "
            "for ($i = 1; $i -le $n; $i++) { "
            "  $p = [System.__ComObject].InvokeMember('Item', $b::GetProperty,"
            "   $null, $cp, @($i)); "
            "  $names += [System.__ComObject].InvokeMember('Name',"
            "   $b::GetProperty, $null, $p, $null) }; "
            "Write-Output ('WORD_PROPS=' + ($names -join ',')); "
            "$d.Save(); $d.Close(); $w.Quit()"
        )
        out = subprocess.run(["powershell.exe", "-NoProfile", "-Command", script],
                             capture_output=True, text=True, timeout=300)
        word_output = (out.stdout or "").replace("\r", "").strip()
        print("\n[2. opened and saved by Microsoft Word]")
        for line in word_output.splitlines():
            print(f"  {line}")
        if out.returncode != 0:
            print("  powershell stderr:", (out.stderr or "").strip()[:400])

        # --- and back to LibreOffice ------------------------------------
        doc = desktop.loadComponentFromURL(docx_linux.as_uri(), "_blank", 0, ())
        after = report(doc, "3. reopened after Word saved it")
        doc.close(False)

        marks_before, bm_before, pages_before = before
        marks_after, bm_after, pages_after = after
        word_sees_prop = "ToaMarks" in word_output
        checks = [
            ("Word itself reads the ToaMarks property", word_sees_prop, True),
            ("custom property survived Word", bool(marks_after), True),
            ("mark count unchanged", len(marks_after), len(marks_before)),
            ("bookmark count unchanged", bm_after, bm_before),
            ("page numbers unchanged", pages_after, pages_before),
        ]
        print(f"\n{'=' * 62}")
        failures = 0
        for name, got, want in checks:
            ok = got == want
            failures += 0 if ok else 1
            print(f"  {'PASS' if ok else 'FAIL'}  {name}"
                  + ("" if ok else f"   got {got!r}, wanted {want!r}"))
        if "WORD_BOOKMARKS=" in word_output:
            print(f"\n  Word's own view: {word_output.splitlines()[0].split('=')[1]}"
                  " bookmark(s), properties: "
                  f"{word_output.splitlines()[1].split('=', 1)[1] or '(none)'}")
        return 1 if failures else 0
    finally:
        proc.terminate()


if __name__ == "__main__":
    sys.exit(main())
