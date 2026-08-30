# Security

## If it involves a real client's brief, do not open an issue

The issue tracker is public and permanent. If reporting something would mean
posting a party's name, a case number, or any part of a real document, email
**Joshua@HarkinsSolutionsSystemsGroup.com** instead.

You do not need a GitHub account for that. If you would rather stay on GitHub,
[private vulnerability reporting][pvr] opens a report only the maintainers can
read.

[pvr]: https://github.com/harkins-solutions/libreoffice-toa/security/advisories/new

Reproducing a bug almost never needs a real document. The extension sorts and
groups the text you hand it and never looks anything up, so an invented
citation in the same form behaves identically to the real one it replaces.

## Reporting a vulnerability

Email **Joshua@HarkinsSolutionsSystemsGroup.com** or use [private vulnerability
reporting][pvr]. Include what you did, what happened, your LibreOffice version
and the extension version. We will confirm receipt, say what we find, and
credit you in the release notes unless you would rather we did not.

Please do not post a working exploit while it is unfixed.

## What this extension does and does not do

| Design | Real finding |
|---|---|
| **No network access.** Nothing is fetched, sent, or checked online. The extension has no URLs in it, and never contacts a citation database. | Any code path that opens a network connection. |
| **Nothing is stored outside the document.** Marks live in the document's own bookmarks and custom properties. No files are written elsewhere, no settings are kept, no telemetry exists. | Anything written outside the document the user has open. |
| **It writes only where asked.** Marking inserts a bookmark over the selection; inserting the table writes a table at the cursor. Nothing else in the document is modified. | Content changed anywhere other than the selection or the cursor. |
| **It runs with your permissions.** A LibreOffice extension is code running as you, with access to your files. That is true of every extension and is why they are installed deliberately. | A path where the extension executes anything from a document — a macro, a script, or content from a `.docx` a third party sent. |

## Marks are visible to anyone who has the file

A mark is a normal bookmark plus a custom document property, both stored in the
document. Anyone with the file can read them — including opposing counsel, if
you send them the working copy.

They carry only what you typed: the authority text and its category. Nothing is
inferred, and no notes, comments or work product are recorded. But if you would
rather the marks not travel with a filed document, generate the table, then
save a copy for filing with the marks removed. Deleting the `ToaMarks` custom
property in **File > Properties > Custom Properties** and the `toa*` bookmarks
leaves the table itself intact — it is ordinary text once written.
