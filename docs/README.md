# Swift NAVFIT user documentation

- [User guide (PDF)](Swift_NAVFIT_User_Guide.pdf): 22 pages with linked contents, chapter bookmarks, numbered procedures, troubleshooting, a practice exercise, and glossary.
- [Online/editable guide](Swift_NAVFIT_User_Guide.md): maintained content for r123a / 0.26.0-dev / schema26.
- [Architecture description](architecture/Swift_NAVFIT_Architecture_Description.md): DoDAF 2.02 fit-for-purpose views in the ISO/IEC/IEEE 42010 structure.
- [Screen-reader test](Accessibility_Screen_Reader_Test.md): the attended NVDA/JAWS procedure.

The supplied eNavFit Version1.0 manual informed the procedure-based organization. The guide describes Swift's local application, not eNavFit's online account or submission services. Orientation tables are screen maps, not captured screenshots.

## Regenerate the PDF

Install Python packages `reportlab` and `pypdf`, then run from the repository:

```powershell
python scripts/build_user_guide.py
```

The renderer uses Segoe UI when available on Windows, otherwise Helvetica. It checks that the 22 chapter/page boundaries are preserved. After changes, render every PDF page and visually inspect pagination, tables, headings, footers, and text. Check the contents links and chapter bookmarks before publishing. Document rendering QA does not replace physical application display/accessibility/printing acceptance.

Version1.0 QA: all22 pages rendered and inspected; 21 chapter bookmarks and20 contents links verified; no blank overflow pages. Application workflows are grounded in r123a source/control definitions and release verification, with older service documentation reconciled to the current navigation.
