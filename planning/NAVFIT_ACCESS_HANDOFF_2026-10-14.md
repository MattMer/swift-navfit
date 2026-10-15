# Access → installed NAVFIT printing handoff investigation

14 October 2026. No Access production code, workspace data or installed NAVFIT settings were changed. No native application was launched, no UI automation was performed and no print job was submitted.

## Conclusion

A visible, assisted handoff is feasible as a design: Swift prepares a persistent NAVFIT print workspace and opens/activates the installed application; the user selects the native folder/report and invokes its existing Print actions. No separate renderer/helper executable or per-job temporary database is required for that design.

A reliable fully automatic database/report/print handoff is not established. Inspected startup paths expose no file/report/print command-line contract, and the assembly is not a VBA COM automation server. Automating File → Open, selection and Print would require a separately tested UI-automation workflow. Do not advertise a one-click Print command based only on launching the application.

## Installed target

Resolved the Start Menu shortcut `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\NAVFIT98\NAVFIT98A.lnk`. It is an MSI-advertised shortcut; its WScript TargetPath resolves to an installer icon executable, not the program to launch. Its working directory identifies the actual installed target:

`C:\Program Files\NIWC Atlantic\NAVFIT98A\NAVFIT98A.exe`

Verified the file exists and SHA-256 is `5366F1C1BB147F1706094C5E3E1D62981F34FC6CB301E4006A892ACFBFAB58CF`, identical to the previously inspected supplied binary. Installed config defaults use ACE.OLEDB.16.0 and the reference catalog filename NF98A_Dte.accdb. Installation presence does not establish Office child-process launch permission on a corporate target computer.

## REA evidence

Used installed REA 6.5.0 CLI, without changing its MCP configuration. Retained complete Evidence in `build/rea-navfit-handoff/startup-members.json`:

`ev_a3acfb404003e8f7582e2f95acc0ba2e42997ec64da11ff90011c42b9b6c3200`

Selected types: `NAVFIT98A.My.MyApplication`, `NAVFIT98A.NF_Main`, `NAVFIT98A.My.MySettings`. REA reports complete selected metadata/CIL coverage. Its external member-reference collection is assembly-wide. The previous [print inspection](NAVFIT_PRINT_REA_2026-10-14.md) provides the PrintModule and COM metadata evidence for the same binary hash/MVID.

[Focused startup/handler evidence](NAVFIT_ACCESS_HANDOFF_REA_2026-10-14.json) records methods, tokens, direct calls and selected startup flags. Flags were additionally decoded from original instruction bytes at REA-reported method/IL locations; that byte read is supplemental shell analysis, not a separate REA producer.

## What the startup and handlers show

- MyApplication declares only Main, its constructor and OnCreateMainForm. Main forwards its argument array to the VB application framework Run method. No custom startup or StartupNextInstance handler is declared on this type.
- The constructor sets IsSingleInstance=True and SaveMySettingsOnExit=True. This supports the expectation that normal launches reuse/activate the application and retain its last settings; runtime activation behavior was not tested here.
- No command-line accessor, print-switch parser, named-pipe server or registered active-object API was found in the selected startup path or relevant assembly-wide external member references. This bounded search does not prove that every undocumented mechanism is impossible. File/report selection through launch arguments is unsupported by the evidence gathered.
- NF_Main_Load (`0x06000bd0`) sets up reference databases and reads saved DBPath. It validates that path, populates the folder tree/report list and enables or disables controls. It does not consume a supplied file argument in this inspected path.
- OpenDatabaseToolStripMenuItem_Click (`0x06000bd4`) constructs an OpenFileDialog titled “Open a Database”, checks the chosen database format, sets DBPath and repopulates the native tree/list. It is a visible file-open workflow, not a public method accepting a path.
- PrintReport_Click (`0x06000beb`) reads tbRptType and selects PrintEval, PrintChief or PrintFitRep. Report data functions subsequently use tbRptId; print correctness depends on native selection state.
- PrintFolder_Click (`0x06000c12`) calls Print_Folder, using native folder selection. PrintSummary_Click (`0x06000c0d`) opens the native summary-letter form rather than directly accepting group data.
- ComVisible(False), previously corroborated by bounded runtime metadata, means the public managed routines do not constitute an ordinary CreateObject/GetObject automation API for VBA.

No supported shell print verb or file-association-based NAVFIT handoff was established. Do not repoint the system .accdb association from Access, invent /print switches, edit installed user.config behind the app, or send guessed window commands as though they were a supported API.

## Recommended first implementation

Label the Swift action **Prepare and open NAVFIT**, with clear handoff status rather than silently claiming a successful print.

1. Save the current report revision. The user chooses a report or summary group in Swift.
2. Write the required native tables into one explicitly configured, persistent `Swift_NAVFIT_Print_Workspace.accdb`. This is separate from the operational Swift workspace and remains on disk. Existing selected report/group exporters supply the native mapping/provenance services, but their destination-overwrite and baseline lifecycle need deliberate adaptation; simply calling them with the same path repeatedly will fail their current overwrite guards.
3. For a first reliable design, require NAVFIT to be closed before refreshing that workspace. An idle native window can retain stale tree/row selections even when no file lock is visible. Preserve the previous print workspace on a failed refresh and never replace files being edited or printed.
4. Launch the installed approved executable from Access VBA, with no unverified path/print arguments and no distributed PowerShell or renderer helper. Discover the installed path instead of hard-coding this computer's location.
5. On first use, or whenever NAVFIT last opened another database, the user selects the persistent print workspace through native File → Open. NAVFIT's saved DBPath can retain it for the next session. Swift displays the exact destination and intended group/member so the user can confirm the selection.
6. The user selects the group/report and uses NAVFIT's Print Report, Print Folder or Print Summary actions. NAVFIT owns its established rendering, printer dialog, page selection and cancellation. Launching the app is not evidence of a completed print; Swift should record “prepared/opened” separately from any user-confirmed “printed” state.

This uses a persistent data exchange database rather than eliminating database exchange. Native populated renderers still query native Reports/Folders/Summary tables; no supported in-memory report-data interface was discovered. Returned edits and any native “validated” flags should not be silently merged back into Swift as part of printing.

## Optional automation, not yet proven

Access could theoretically operate native controls through accessibility/Win32 mechanisms. It would need to bind the exact process/window, confirm the database, select stable report/folder identities, wait for native population, detect modal errors and stop at the user's print dialog. Focus-based SendKeys, fixed screen coordinates and blind delays would be fragile. An approved executable being present is not proof that those automation mechanisms or Office launching it are permitted on the deployment environment.

No such automation was attempted, in line with the preference to avoid computer use. If a fully automatic handoff is essential, the next investigation should explicitly test those native interactions on a synthetic persistent workspace. Otherwise implement the assisted workflow above first.

## Proof still required before rollout

Actual Access-to-NAVFIT launch/activation on the target corporate machine; first-use and remembered-workspace opening; refresh/reopen safety; single report, group and summary selection; cancel/error handling; actual native print output and correct report identity. Current investigation establishes static interfaces and installed target identity, not end-to-end runtime printing.
