# NAVFIT98A print handler — REA inspection and reuse proof

14 October 2026. This is an investigation and separate runtime proof; Swift NAVFIT's Access application was not changed and no release was built or published.

## Result

There is a credible path to reuse NAVFIT98A's actual report drawing code through an out-of-process .NET helper. Direct invocation of five drawing-only methods successfully produced EVAL/FITREP page-two frames and labels plus a CHIEFEVAL page-two frame. This proves reuse of those methods, not the complete populated print workflow.

Recommend a bounded populated-EVAL proof next, using Swift's existing Patch 2 export as input. If full native state initialization proves reliable, the helper can become Swift's print/preview backend. Keep rendering and fit assessment in that same helper so the editor does not approximate one engine while printing through another.

## Target and REA evidence

Inspected the supplied `NAVFIT98A/NAVFIT98A.exe`: assembly 2.2.0.332, .NET Framework 4.8.1, x64 managed CIL. SHA-256 `5366F1C1BB147F1706094C5E3E1D62981F34FC6CB301E4006A892ACFBFAB58CF`; MVID `74a3c3b5-a7af-48bc-ac92-963eb5331ff0`.

REA 6.5.0 is available in the local npm cache. Doctor found no Codex MCP registration and this chat exposes no REA MCP tools. Used the installed CLI directly; no configuration was changed and no native decompiler/engine was installed. REA's managed inspection reads file-backed metadata/CIL without executing the target. Runtime proof below is separate ordinary PowerShell/.NET work, not REA runtime evidence.

| Scope | REA Evidence ID |
|---|---|
| Assembly identity, metadata and attributes | `ev_3e422ae457e0fd5eb6ad3660ca7cb477afaa40a34fea8af201a78dc49cfc49e0` |
| PrintModule: 63 methods, signatures, CIL anchors, calls and field accesses | `ev_46aad1dfad11896871de7ed0a1c6309a4262c716bedd614752518ade0402ab18` |
| Settings, calculations, address formatting and reference-data loading | `ev_b8e69fb73ba50e53e95dff843d978efd5b4a540166103aeb1f8c4e9a69df3ef1` |

[Focused findings and exact method tokens](NAVFIT_PRINT_REA_2026-10-14.json). Complete original CLI Evidence JSON is retained locally in `build/rea-navfit-print/{artifact,print-members,dependencies}.json`; `narrative-anchors.txt` contains readable selected CIL anchors. Tokens are valid only with this exact digest/MVID. Selected method coverage is complete as reported by REA, but it is not a whole-application behavioral proof or recovered original C#/VB source.

## Print flow observed in static evidence

- Public static entry methods `PrintEval` (`0x06000cd2`), `PrintChief` (`0x06000cce`) and `PrintFitRep` (`0x06000ccf`) create PrintDocument/PrintDialog objects, attach their PrintPage callbacks, display the dialog and call Print(). They do not accept report-data arguments.
- `NF_PrtEval` (`0x06000cd4`) selects page-one/page-two text, frame and data functions and changes HasMorePages. The module uses a shared PrintPage field. Chief and FITREP have corresponding paths; folder printing also uses report/folder counters.
- Frames and most EVAL/FITREP fixed wording use Graphics.DrawLines/DrawRectangle/DrawString. CHIEFEVAL's text functions also load trait descriptions from the reference catalog via LoadTraitBlocks; drawing all labels is not database-free for that form.
- Page data functions open an OleDb connection and get the native report ID through `My.Forms.NF_Main.tbRptId`. Group calculations read `tbFolderSelection` and write trait sum/count controls. Address text passes through BuildRSAddressToPrint rather than simply concatenating Swift controls.
- Folder/summary methods query native data, rely on selected-folder state, and use separate pagination. Summary printing additionally uses NF_Summary controls. Implementing only individual pages would leave this workflow unfinished.
- DBPath, DBACCDBPathRef, DBACCDBPathDte and DBProvider settings supply report/catalog connections. The supplied catalog exists at `NAVFIT98A/Data/NF98A_Dte.accdb`. A helper must supply explicit paths rather than inherit a stale root-drive path.
- Runtime metadata corroborates assembly `ComVisible(False)`. Public managed methods are not a ready-made VBA COM API. A separate executable avoids in-process Access CLR/bitness coupling; its own ACE/.NET dependencies still require target-computer verification.

## Why a matching text box is insufficient

REA's CIL shows EVAL narrative selection by CommentsPitch. The default branch uses UseCNFont9; a stored pitch starting with `12` selects UseCNFont10 and a different rectangle. The latter path splits words, measures candidate lines using Graphics.MeasureString, reads Font.GetHeight and uses a multiplier about 0.82 for line advance. FITREP/CHIEFEVAL have their own drawing paths. The module also defines a Courier New 9.65-point font used for some trait text.

The fixed Access editor profile from r104 remains an input convention. It does not reproduce these renderers or establish printed line capacity. A reusable print/measurement helper is the stronger route for narrative preview and overflow feedback.

## Bounded runtime proof

Ran `scripts/prove-navfit-drawing.ps1` under Windows PowerShell in a separate process. It hash-gates this exact executable, loads it, creates 850×1100-pixel bitmaps at 100 DPI and invokes only:

- NF_EvalPage2Frame + NF_EvalPage2Text;
- NF_FitRepPage2Frame + NF_FitRepPage2Text;
- NF_ChiefPage2Frame.

REA's selected direct calls showed these methods and the PrintModule initializer use drawing/font/string-format operations, without the report-data/UI/database paths. Printed runtime metadata also recorded actual font sizes and COM visibility. All three outputs were nonempty. The EVAL image was visually inspected. Input binary hash remained unchanged. No application entry point, print dialog, Print(), report data method, CHIEF text method or physical printer job was invoked.

[Runtime record and image hashes](NAVFIT_PRINT_RUNTIME_PROOF_2026-10-14.json). Local output: `output/NAVFIT_Print_Proof_2026-10-14/`, including `EVAL-page2.png`, `FITREP-page2.png`, `CHIEFEVAL-page2-frame.png`, and `proof.json`.

These are blank report surfaces on bitmap Graphics, not populated reports or proof of printer geometry, margins, pagination or clipping. Reusing the same methods on a printer Graphics still needs comparison with NAVFIT's actual output.

## Proposed Swift implementation

1. Swift commits the draft and exports a disposable Patch 2 database using its existing report/group exporter. Pass native exported report/folder IDs explicitly; do not substitute Swift's internal IDs.
2. Launch a small x64 .NET Framework helper with installed NAVFIT98A path, exported database/catalog paths, report or group selection, operation (preview/print), and output destination. Gate supported versions/digests; keep each job isolated. Do not register or embed the assembly in Access.
3. Initialize required native settings, hidden form controls and page counters only inside that owned process. Use a read-only staged database; never attach directly to the live Swift workspace. Prove that hidden form initialization does not depend on additional application startup or mutate installed user settings.
4. Drive the native page callbacks through a controlled PrintDocument and preview controller. Use the same renderer for fit/overflow analysis and final printing. Actual printing remains a user-triggered action. Native summary-group arithmetic and Swift's intended group membership need explicit comparison.
5. Return a structured success/error manifest to Access, with output paths and selected report identity. A failed helper cannot lock the Access editor or alter saved report data.

This approach uses the installed NAVFIT runtime as a version-specific dependency. A later independent renderer could adapt the drawing/data routines and eliminate hidden native UI state, but would require a larger comparison matrix. Simply translating them into Access reports is unlikely to preserve GDI+ wrapping exactly.

## Remaining proof before implementation is ready

First render both pages of a populated synthetic EVAL and compare with a native NAVFIT PDF of the same exported record. Then repeat for FITREP, CHIEFEVAL and summary letters, both pitch settings, NOB, blank values, multiple paragraph styles, Unicode, long words and exact boundary-length narratives. Verify group selection, native-only values, arithmetic, printer margins, page size/orientation and cancellation. Test the helper on a second target computer with the expected .NET/ACE installation.

No full populated rendering, automated PDF preview, physical print, full native form initialization or production integration is claimed by this investigation.

## Reproduce

`scripts/inspect-navfit-print-rea.ps1` accepts the explicit installed REA entry script, Node executable, supplied target and a new analysis directory. It captures complete Evidence without modifying registration. `scripts/summarize-navfit-print-rea.cjs` builds the focused report and CIL listing. `scripts/prove-navfit-drawing.ps1` performs the bounded runtime proof and refuses another binary hash or existing output directory.
