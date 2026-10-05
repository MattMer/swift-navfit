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
