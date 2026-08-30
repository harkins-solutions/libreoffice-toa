"""Table of Authorities for LibreOffice Writer.

A mark is two things stored in the document itself:

  * a bookmark named `toa<n>` covering the cited text, which supplies the
    page number and moves with the text when the brief is edited;
  * an entry in a JSON blob held in the user-defined document property
    `ToaMarks`, which supplies the authority string and its category.

Both survive .odt and .docx round trips, so the table can be rebuilt after any
edit and the page numbers will be current.

Dispatch URLs (also on the Authorities menu):

  org.libreoffice.toa:MarkCitation   mark the selection; with Authority= and
                                     Category= arguments it skips the dialog
  org.libreoffice.toa:InsertTable    build the table at the cursor
  org.libreoffice.toa:ListMarks      report what is marked
"""
import json

import unohelper
from com.sun.star.awt import Rectangle
from com.sun.star.frame import XDispatch, XDispatchProvider
from com.sun.star.lang import XInitialization, XServiceInfo

IMPL_NAME = "org.libreoffice.toa.Handler"
PROTOCOL = "org.libreoffice.toa:"
PROP_NAME = "ToaMarks"
BOOKMARK_PREFIX = "toa"

# Order they appear in the finished table.
CATEGORIES = ("Cases", "Statutes", "Rules", "Constitutional Provisions", "Other")
PASSIM_AT = 5


# -- the mark store ---------------------------------------------------------

def read_marks(doc):
    props = doc.getDocumentProperties().getUserDefinedProperties()
    try:
        blob = props.getPropertyValue(PROP_NAME)
    except Exception:
        return {}
    return json.loads(blob) if blob else {}


def write_marks(doc, marks):
    props = doc.getDocumentProperties().getUserDefinedProperties()
    blob = json.dumps(marks)
    try:
        props.setPropertyValue(PROP_NAME, blob)
    except Exception:
        props.addProperty(PROP_NAME, 0, blob)


def _next_name(marks):
    used = {int(n[len(BOOKMARK_PREFIX):]) for n in marks
            if n.startswith(BOOKMARK_PREFIX) and n[len(BOOKMARK_PREFIX):].isdigit()}
    return f"{BOOKMARK_PREFIX}{(max(used) + 1) if used else 0:04d}"


# -- overlap detection ------------------------------------------------------
#
# compareRegionEnds(a, b.getStart()) == 1 means a ends before b starts, i.e.
# they are disjoint. Measured rather than assumed; see tests/test_overlap.py.

def _disjoint(text, a, b):
    try:
        return (text.compareRegionEnds(a, b.getStart()) == 1
                or text.compareRegionEnds(b, a.getStart()) == 1)
    except Exception:
        # Different text objects (a header, a frame) cannot overlap the body.
        return True


def existing_mark_at(doc, text_range):
    """The name of a mark already covering this range, or None.

    Without this, marking the same citation twice silently accumulates
    bookmarks: the table still looks right, because entries are grouped by
    authority, while the document fills with duplicates that later edits and
    deletions then disagree about.
    """
    text = doc.getText()
    bookmarks = doc.getBookmarks()
    for name in read_marks(doc):
        if not bookmarks.hasByName(name):
            continue
        anchor = bookmarks.getByName(name).getAnchor()
        if not _disjoint(text, anchor, text_range):
            return name
    return None


# -- marking ----------------------------------------------------------------

def guess_category(authority):
    lowered = authority.lower()
    if " r. " in lowered or "rule" in lowered:
        return "Rules"
    if "const." in lowered or "amend." in lowered:
        return "Constitutional Provisions"
    if "stat" in lowered or "u.s.c" in lowered or "code" in lowered or "§" in authority:
        return "Statutes"
    return "Cases"


def mark_range(doc, text_range, authority, category):
    """Record one authority. Returns (name, status).

    status is 'marked', 'updated' when the range was already marked and the
    category changed, or 'duplicate' when it was already marked identically.
    """
    authority = (authority or text_range.getString()).strip()
    if not authority:
        return None, "empty"
    category = category if category in CATEGORIES else guess_category(authority)

    marks = read_marks(doc)
    existing = existing_mark_at(doc, text_range)
    if existing is not None:
        current = marks.get(existing, {})
        if current.get("authority") == authority and current.get("category") == category:
            return existing, "duplicate"
        marks[existing] = {"authority": authority, "category": category}
        write_marks(doc, marks)
        return existing, "updated"

    name = _next_name(marks)
    bookmark = doc.createInstance("com.sun.star.text.Bookmark")
    bookmark.setName(name)
    doc.getText().insertTextContent(text_range, bookmark, True)
    marks[name] = {"authority": authority, "category": category}
    write_marks(doc, marks)
    return name, "marked"


# -- building the table -----------------------------------------------------

def page_of(doc, text_range):
    cursor = doc.getCurrentController().getViewCursor()
    cursor.gotoRange(text_range, False)
    return cursor.getPage()


def _pages_label(pages):
    """'3', '3, 7', '3-5, 9', or 'passim' past the threshold."""
    ordered = sorted(pages)
    if len(ordered) >= PASSIM_AT:
        return "passim"
    runs = []
    start = previous = ordered[0]
    for page in ordered[1:]:
        if page == previous + 1:
            previous = page
            continue
        runs.append((start, previous))
        start = previous = page
    runs.append((start, previous))
    return ", ".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def collect(doc):
    """{category: [(authority, pages label)]}, alphabetised within category."""
    marks = read_marks(doc)
    bookmarks = doc.getBookmarks()
    pages = {}
    for name, meta in marks.items():
        if not bookmarks.hasByName(name):
            continue          # the bookmark went with the text
        anchor = bookmarks.getByName(name).getAnchor()
        if not anchor.getString().strip():
            # Deleting the cited text leaves the bookmark behind, collapsed to
            # a point. Listing an authority the brief no longer cites would
            # misstate the record, so an empty anchor is not an entry.
            continue
        page = page_of(doc, anchor)
        pages.setdefault((meta["category"], meta["authority"]), set()).add(page)

    grouped = {}
    for (category, authority), page_set in pages.items():
        grouped.setdefault(category, []).append(
            (authority, _pages_label(page_set)))
    for entries in grouped.values():
        entries.sort(key=lambda e: e[0].lower())
    return grouped


def insert_table(doc):
    grouped = collect(doc)
    if not grouped:
        return 0
    text = doc.getText()
    cursor = doc.getCurrentController().getViewCursor()

    rows = 1 + sum(1 + len(grouped[c]) for c in CATEGORIES if c in grouped)
    table = doc.createInstance("com.sun.star.text.TextTable")
    table.initialize(rows, 2)
    text.insertTextContent(cursor, table, False)

    table.getCellByName("A1").setString("TABLE OF AUTHORITIES")
    table.getCellByName("B1").setString("Page(s)")
    row = 2
    for category in CATEGORIES:
        if category not in grouped:
            continue
        table.getCellByName(f"A{row}").setString(category.upper())
        row += 1
        for authority, label in grouped[category]:
            table.getCellByName(f"A{row}").setString(authority)
            table.getCellByName(f"B{row}").setString(label)
            row += 1
    return rows


# -- the dialog -------------------------------------------------------------

def ask_for_mark(ctx, default_authority, default_category):
    """Ask for the authority text and its category. None if cancelled."""
    smgr = ctx.ServiceManager
    model = smgr.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialogModel", ctx)
    model.Width, model.Height, model.Title = 240, 105, "Mark Citation"

    def add(kind, name, **props):
        control = model.createInstance(f"com.sun.star.awt.UnoControl{kind}Model")
        for key, value in props.items():
            setattr(control, key, value)
        model.insertByName(name, control)
        return control

    add("FixedText", "lblAuth", PositionX=8, PositionY=8, Width=60, Height=12,
        Label="Authority:")
    add("Edit", "txtAuth", PositionX=70, PositionY=6, Width=162, Height=14,
        Text=default_authority)
    add("FixedText", "lblCat", PositionX=8, PositionY=30, Width=60, Height=12,
        Label="Category:")
    combo = add("ListBox", "lstCat", PositionX=70, PositionY=28, Width=162,
                Height=14, Dropdown=True, StringItemList=tuple(CATEGORIES))
    combo.SelectedItems = (CATEGORIES.index(default_category),)
    add("FixedText", "lblHint", PositionX=8, PositionY=50, Width=224, Height=24,
        Label=("Short forms and id. cites pointing at this authority should be "
               "marked too, with the same text."), MultiLine=True)
    add("Button", "btnOk", PositionX=126, PositionY=82, Width=50, Height=16,
        Label="Mark", PushButtonType=1)
    add("Button", "btnCancel", PositionX=182, PositionY=82, Width=50, Height=16,
        Label="Cancel", PushButtonType=2)

    dialog = smgr.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialog", ctx)
    dialog.setModel(model)
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    dialog.createPeer(toolkit, None)
    try:
        if dialog.execute() != 1:
            return None
        return (dialog.getControl("txtAuth").getModel().Text,
                CATEGORIES[dialog.getControl("lstCat").getModel().SelectedItems[0]])
    finally:
        dialog.dispose()


def _message(ctx, frame, text, title="Table of Authorities"):
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    toolkit = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.Toolkit", ctx)
    parent = frame.getContainerWindow() if frame else None
    box = toolkit.createMessageBox(parent, 1, BUTTONS_OK, title, text)
    box.execute()
    box.dispose()


# -- the dispatch handler ---------------------------------------------------

def _args(properties):
    return {p.Name: p.Value for p in (properties or ())}


class Handler(unohelper.Base, XServiceInfo, XDispatchProvider, XDispatch,
              XInitialization):
    def __init__(self, ctx):
        self.ctx = ctx
        self.frame = None

    def initialize(self, args):
        if args:
            self.frame = args[0]

    def getImplementationName(self):
        return IMPL_NAME

    def supportsService(self, name):
        return name == "com.sun.star.frame.ProtocolHandler"

    def getSupportedServiceNames(self):
        return ("com.sun.star.frame.ProtocolHandler",)

    def queryDispatch(self, url, target_frame_name, search_flags):
        return self if url.Protocol == PROTOCOL else None

    def queryDispatches(self, requests):
        return tuple(self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
                     for r in requests)

    def dispatch(self, url, properties):
        doc = self.frame.getController().getModel() if self.frame else None
        if doc is None:
            return
        args = _args(properties)
        if url.Path == "MarkCitation":
            self._mark(doc, args)
        elif url.Path == "InsertTable":
            insert_table(doc)
        elif url.Path == "ListMarks":
            self._list(doc, args)

    def _mark(self, doc, args):
        selection = doc.getCurrentController().getSelection()
        if not selection.getCount():
            return
        text_range = selection.getByIndex(0)
        authority = args.get("Authority") or text_range.getString().strip()
        category = args.get("Category") or guess_category(authority)

        if not args.get("Authority"):        # interactive: ask
            answer = ask_for_mark(self.ctx, authority, category)
            if answer is None:
                return
            authority, category = answer

        name, status = mark_range(doc, text_range, authority, category)
        if status == "duplicate" and not args:
            _message(self.ctx, self.frame,
                     "That text is already marked as this authority.")

    def _list(self, doc, args):
        grouped = collect(doc)
        lines = []
        for category in CATEGORIES:
            for authority, label in grouped.get(category, []):
                lines.append(f"{category}: {authority} — {label}")
        text = "\n".join(lines) or "Nothing is marked yet."
        if not args:
            _message(self.ctx, self.frame, text, "Marked Citations")
        return text

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Handler, IMPL_NAME, ("com.sun.star.frame.ProtocolHandler",))
