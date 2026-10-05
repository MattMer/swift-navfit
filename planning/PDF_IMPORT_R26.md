# PDF import expansion and existing-draft merge — r26, 4 October 2026

This implements the user's clarified priority #4: expand the eight-field PDF
mapping, merge into existing drafts, and verify the extraction dependency on
target computers. It does not refer to the older release-packaging priority #4.
Version remains 0.22.0-ux-preview, schema 21; no data schema change.

## Workflow

Reports > Import reports > Choose structured PDF:

1. Choose a completed supported May-2025 interactive NAVPERS PDF.
2. Choose the explicit century for six-digit dates. SSN is the default.
   Legacy DoDID remains available only with deliberate compatibility selection.
3. Leave the target blank for a new draft, or select an existing editable draft
   of the same type. Select **Fill empty fields** or **Replace mapped values**.
4. Preview the full existing, incoming and resulting values. Filled fields that
   are preserved are explicitly shown. Missing/blank source values never clear
   existing fields. Apply is enabled only after a current nonempty preview.
5. Apply creates normal atomic revisions and a receipt with source XML, template
   identity, file hash, applied revision and (for merges) preceding revision.
   Results opens the affected report. Existing local notes, display names and
   separately stored name parts are preserved by a merge.

Changed source, mapping choices or target revisions require another preview.
Archived/approved drafts, different form types, an open target editor or pending
target recovery prevent merging. The existing Operation history import undo
restores a merged draft through a compensating revision; later edits/locks block
undo. Exact duplicate PDFs are blocked across the old and new PDF adapter versions.

## Reviewed mapping scope

The source catalogue is `src/pdf-field-map.json`, emitted with
`scripts/generate-pdf-map.py` into `modPDFMapping.bas`. The read-only Python
preview uses the same catalogue and retains raw values without pretending they
are already normalized native draft values.

| Saved-field coverage | EVAL | CHIEFEVAL | FITREP |
|---|---:|---:|---:|
| Catalogue entries | 40 | 51 | 36 |

All three include official name/rate/designator/SSN, member UIC/ship-station,
component, promotion status, reporting period/date reported, six reporting-senior
identity fields, primary duty/description, counseling date/name, seven traits,
two assignment recommendations, billet subcategory, promotion recommendation
and official comments. EVAL additionally includes rater/senior-rater names/dates.
CHIEFEVAL additionally includes named occasion/type/NOB/statement checkboxes,
readiness date and individual address lines/commercial phone/DSN.

Bindings were checked against the supplied templates' XFA field/exclusive-group
names, captions, tooltips, block positions and saved item values. EVAL/FITREP
traits save `nob` or `1`–`5`; CHIEFEVAL dropdowns save `0`–`5`. EVAL promotion
status saves `1`–`4`. Component saves `act/tar/inact/at`; native explicit Yes/No
values are derived from that selection. Named CHIEFEVAL checkboxes save
`1/0/2` (selected/unselected/neutral); neutral is omitted. Unsupported codes are
rejected, never guessed. Date values require the explicit century and real-date
checks. Length limits are enforced by the normal native draft commit without
truncation; an overlong source requires deliberate review/correction.

**Partial mapping remains:** EVAL/FITREP unnamed occasion/type/NOB/retention/
statement checkbox bindings need further controlled native-save proof. Command
employment/achievements and EVAL qualifications have no dedicated saved target.
Combined address fields are not split heuristically. Combined physical readiness
text, signatures, concurrent senior/signature fields, supplied ranking, RSCA,
summary counts and calculated averages are retained as source XML, not imported
as verified draft values. No inferred signatures/approvals or schema extension.
Flattened/scanned/general AcroForm PDFs and other template revisions are rejected.
Printing/rendered fit remains deferred; synthetic XFA fixtures test data extraction.

## Extraction dependency and workstation check

This adapter requires desktop Adobe Acrobat automation:
`AcroExch.PDDoc.Open` > `GetJSObject` > XFA template/datasets `saveXML`.
Adobe Reader's limited OLE/browser control is insufficient; see
[Adobe's official IAC documentation](https://opensource.adobe.com/dc-acrobat-sdk-docs/library/interapp/IAC_DevApp_OLE_Support.html).

The import screen reports automation availability without claiming successful
extraction. After choosing a PDF, **Check PDF extraction on this computer**
actually reads its XFA packets and checks that the source bytes stayed unchanged.
An actionable failure identifies the required dependency; no installation,
licensing, trust or security settings are changed automatically.

For a repeatable target-computer acceptance check, use the portable dependency
kit with the r26 app and run:

```powershell
powershell.exe -NoProfile -File .\scripts\test-pdf-target-computer.ps1 `
  -BuildPath 'C:\YourFolder\Swift_NAVFIT_UI_Demo_2026-10-04-r26.accdb'
```

The checker uses a temporary copy, records the actual Access build/bitness and
Acrobat viewer version, verifies extraction/template identity/counts/selected
expected synthetic values for all three forms, and writes a timestamped JSON
report. The supplied app and PDFs must remain byte-identical. Run as the intended
user under normal workstation settings. The script requires desktop Access COM
and module import; the native check button can exercise an installed build
without importing the harness. Runtime-specific and x86 certification remain
separate acceptance work. No claim is made for a computer that has not run it.

## Verification

Final local artifacts:

- Demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r26.accdb` — 9 synthetic reports,
  3 groups, 2 contributors, 1 senior, 3 RSCA baselines and review/approval/archive/
  NOB/recovery examples. SHA256
  `31BA535B33D0D8CBAF008B70A4BFE336EF4862139198E8DC8B0CE33933144C44`.
- Blank: `build/ux-r26-final.accdb` — all 42 user-data tables empty. SHA256
  `DCAB6B8FF589D1A698A4D5C06F5D5700C5CE7269BA42E567DC9D53C5DDA7A1D1`.
- Portable checker: `output/pdf/Swift_NAVFIT_PDF_Dependency_Check_r26.zip`.
  SHA256 `9054357BE3C76A59436C2C72916AC52546E005E6DA8D02200DC985E132A8B370`.

Verification passed on disposable copies of the exact final blank bytes:

- **250** expanded extraction/field decoding/commit/merge/undo assertions;
  **28** original PDF assertions; **43** guided import assertions, including
  changed source/mapping choices and repeated Apply. Expanded cases include all
  mapped saved values for each form, null/blank/NOB/malformed traits, stale target,
  changed mode, wrong form, archive/approval locks and pending recovery.
- **145** review/Reports return, **59** live validation/source, **32** My Work,
  **30** help consistency, **23** calculation/checklist, **41** group lifecycle,
  **90** blocks/defaults/comparison, **1793** responsive, **108** editor surface,
  **49** native workflow assertions and two forced-shutdown recovery cases.
- **36** populated demo assertions plus normal Office and enabled-content startup.
- All **53** installed standard modules match source in both final files; native
  review/return hooks and compilation verified. Deliverable hashes unchanged.
  Evidence: `build/tests/review-build-67522de2-0c4a-4a30-be3b-ae4e2dab609d/verification.json`
  (blank) and `build/tests/review-build-0605075f-6678-456a-9f25-707e244039ae/verification.json`
  (demo). Blank table counts: `build/tests/r26-blank-counts.json`.

Local dependency result: Windows build 26200, **Access 16.0 build 20430 64-bit**,
**Acrobat viewer version 26.00221931**, all three expanded synthetic form probes
PASS. Report: `output/pdf/dependency-20261004-223623.json`. The actual ZIP was
unpacked and its checker also passed against the final demo, without changing
the app or PDFs. This proves the kit's local portability, not untested computers.

Native visual check on `build/tests/r26-pdf-visible.accdb` showed the target and
mode selectors, complete before/PDF/after preview, Apply/Results and dependency
button without clipping at the current window size. A Sky click on the native
dependency button displayed successful XFA extraction and the Acrobat version.
The probe was closed; this is not a six-configuration physical DPI or screen-reader
acceptance claim. Source XFA fixture bytes stayed unchanged during native tests.
No PDF print appearance claim is made.

Earlier files and existing workspaces remain intact; no GitHub publication is
part of this change. Runtime/x86, other approved workstations and the remaining
unsupported PDF fields need further acceptance/implementation.
