# Swift NAVFIT user documentation

- [User guide (PDF)](Swift_NAVFIT_User_Guide.pdf): 38 pages with linked contents, screenshots, chapter bookmarks, numbered procedures, troubleshooting, a practice exercise, and glossary. Version 1.6 for Swift NAVFIT 0.52.0, schema 41.
- [Online/editable guide](Swift_NAVFIT_User_Guide.md): the maintained source of the PDF.
- [Architecture description](architecture/Swift_NAVFIT_Architecture_Description.md): DoDAF 2.02 fit-for-purpose views in the ISO/IEC/IEEE 42010 structure.
- [Screen-reader test](Accessibility_Screen_Reader_Test.md): the attended NVDA/JAWS procedure.

The supplied eNavFit Version1.0 manual informed the procedure-based organization. The guide describes Swift's local application, not eNavFit's online account or submission services. Orientation tables are screen maps, not captured screenshots.

## Regenerate the PDF

Install Python packages `reportlab` and `pypdf`, then run from the repository:

```powershell
python scripts/build_user_guide.py
```

The renderer uses Segoe UI when available on Windows, otherwise Helvetica. Each chapter starts on a new page and may run over several. A first pass measures where chapters start, and the second fills those page numbers into the contents table. The renderer reads the version line from the guide itself. After changes, render every PDF page and visually inspect pagination, tables, headings, bullets, footers, and text. Check the contents links and chapter bookmarks before publishing. Document rendering QA does not replace physical application display/accessibility/printing acceptance.

Version 1.3 QA (26 October 2026): rendered to 25 pages; contents page numbers, chapter starts, bullets and tables checked visually.
