# Installed NAVFIT98A print engine inspection

Inspected 5 October 2026: supplied and installed NAVFIT98A 2.2.0.332. This is
metadata/IL evidence from the local executable, not a print-output certification.

The supplied executable is a managed .NET Framework application. Its relevant
type and method names remain readable. `NAVFIT98A.PrintModule` exposes public
static methods such as `PrintEval`, `PrintChief`, `PrintFitRep`, `Print_Folder`,
`PrintSummary`, and per-page text/frame/data drawing routines. These use
`System.Drawing.Printing.PrintDocument`, `PrintPageEventArgs`, `Graphics`, fonts,
and drawing operations. They are code-based layouts rather than embedded PDF
report templates. No obfuscation prevents inspection of these particular paths;
this does not establish that every part of the executable is unobfuscated.

"Public" here means accessible to another .NET assembly. The assembly has
`ComVisible(False)`, so these methods are not an Access VBA COM automation API.
The print paths depend on VB `My.Forms`, the main form's selected folder/report,
`My.Settings.DBPath`, database connections and module-level page counters. Folder
queries match the report parent to `"a " + selected FolderID`; summary printing
also reads the summary-letter form. A simple call from Access would not supply
that state.

Using NAVFIT98A itself to open Swift's exported ACCDB and print uses its actual
engine. Printer driver, paper and settings still affect output. A separate .NET
helper could potentially load the installed assembly, initialize its state and
invoke these routines, or adapt the drawing routines. Neither path is implemented
or verified. Direct reuse would be version-coupled; adapting layouts would
require visual comparison of all report pages and summary letters. No independent
command-line/headless print contract was established by this inspection.

Reproduce inspection with `scripts/inspect-navfit-print-export.ps1`. Generated
method evidence is in `planning/evidence/navfit/native-print-export.json`.
No target print method was executed and no physical print job was submitted.

## Narrative-control measurements - 6 October 2026

`inspect-navfit-narrative-layout.ps1` disassembles 229 methods, including native form initialization and per-page drawing helpers. It decodes floating-point IL constants; no NAVFIT methods run. Evidence: `planning/evidence/navfit/native-narrative-layout.json`, assembly version 2.2.0.332 and SHA256 5366F1C1BB147F1706094C5E3E1D62981F34FC6CB301E4006A892ACFBFAB58CF.

All three form designer controls initialize Courier New at 12 points, multiline:

| Native control | Designer size (pixels) | Designer MaxLength |
|---|---|---|
| EVAL tbEComments | 921 x 386 | 2500 |
| CHIEF tbCComments | 883 x 133 | 700 |
| FITREP tbFComments | 945 x 378 | 2000 |

These are designer defaults, not authoritative print-fit capacities or runtime scaled dimensions. EVAL and FITREP print routines branch on stored CommentsPitch; the 12-pitch branch selects UseCNFont10. Copying the designer's 12-point font into Swift does not prove equivalent printed wrapping. Runtime pitch changes, text wrapping, drawing rectangles, margins and printer output still need verification before item 23 is complete. No new Swift font/length restriction has been applied from this partial evidence.


## Runtime pitch and print branches - r67 investigation

The additional 237-method inspection includes pitch-change handlers and the
PrintModule type initializer. EVAL/FITREP selectors contain `10 Point` then
`12 Point`. Their editing handlers use Courier New 12pt at 925 x 330 pixels for
the first choice, and Courier New 14pt at 920 x 390 for the second. These differ
from designer initialization and from printing; preserving this distinction is
necessary when implementing a locked editor.

PrintModule initializes UseCNFont9 at 9pt and UseCNFont10 at 10pt. The narrative
branches observed in IL are:

| Type | Default print path | Stored pitch starting with 12 |
|---|---|---|
| EVAL | Courier New 9pt, rectangle x35/y458/w715/h260 | Courier New 10pt, x35/y468/w715/h275; custom word-measured lines |
| FITREP | Courier New 9pt, rectangle x35/y478/w715/h260 | Courier New 10pt, x35/y478/w715/h295; DrawString rectangle |
| CHIEFEVAL | Courier New 10pt, rectangle x26/y590/w760/h115 | No pitch branch observed in this comments path |

Coordinates are raw drawing constants; conversion to physical units still
requires verification of Graphics transforms and page/printer settings. EVAL's
12-point branch splits on space, measures candidate lines with Graphics.MeasureString,
uses Font.GetHeight(Graphics) for the clipping condition, and advances lines by
0.82 of that height. FITREP and CHIEFEVAL use rectangle-based DrawString instead.
An Access multiline text box uses a different rendering engine. A fixed font or
character count therefore cannot certify the same wrapping/clipping.

No target print method was executed, no printer settings were changed, and no
physical print job was submitted. Item 23 remains incomplete until a measured
preview and comparisons verify these paths. The supplied executable's current
code is the reference, rather than guessed line/character limits.
