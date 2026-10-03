# NAVFIT sample findings

Inspected 3 October 2026. User identifies `TEST_DB.accdb` as NAVFIT98A v33 Patch 2. Exact executable version/build and reimport/printing are not yet verified. The sample is evidence of this file's schema, not a support claim for every database labeled v33 Patch 2.

Original SHA256: `E732EE41A384495AC37BC009C59516A73461C7DB99E251B5384CE917652386E7`. Read-only DAO inspection used a scratch copy and verified the original hash remained unchanged.

| Table | Columns | Rows | Notes |
|---|---:|---:|---|
| Reports | 128 | 4 | ReportID Long Integer primary key; Parent Short Text(12); ReportType Short Text(15) |
| Folders | 78 | 4 | FolderID Long Integer; Parent Integer; includes inherited-looking report fields whose behavior is unverified |
| Summary | 6 | 0 | ProfileName, OnTimeYes/No, ISICAddress, EMail, Phone; do not infer that this table stores official summary-group membership |

Reports has one `Eval`, one `Chief`, one `FitRep` and one null ReportType. Unknown/null types must be quarantined in staging and explained, not assigned a guessed form. No personal field values were exported during inspection. Actual fixture data is not yet certified synthetic.

Report and folder parent types differ. Referential enforcement was not exposed through DAO Relations; do not infer folder joins or roots solely from field names. Required fields, indexes, defaults, data types and sizes are recorded in `evidence/navfit/schema-inventory.json`. ReportID is an Integer-sized Long key, not a Swift GUID. Source IDs must be scoped to source database identity; duplicate IDs across files do not prove duplicate reports.

Verified external columns relevant to candidate local mapping:

| Swift concept | Observed external column | Constraint | Mapping status |
|---|---|---|---|
| Report form kind | Reports.ReportType | Text(15), observed Eval/Chief/FitRep/null | Candidate enum map; null rejected; actual NAVFIT behavior pending |
| Member display identity | Reports.FullName; FirstName/MI/LastName/Suffix | Text(27/27/3/27/10) | Need decomposition/normalization proof; not enough for identity matching |
| Period start/end | Reports.FromDate/ToDate | Date/Time | Candidate; blank and NAVFIT normalization tests pending |
| Report narrative | Reports.Comments | Long Text | Do NOT map Swift DraftNotes automatically; official form/block and layout proof pending |
| Related folders | Reports.Parent, Folders.FolderID/Parent | Text(12), Long, Integer respectively | Logical relation unresolved |
| Traits | PROF/QUAL/EO/MIL/PA/TEAM/LEAD/MIS/TAC and associated DN1/DN2/DN3 fields | Integer marks; DN fields Text(255) | Form-specific meaning/code map unresolved |
| Promotion recommendation | Reports.PromotionRecom | Integer | Numeric enum values not interpreted yet |
| RSCA | Reports.RSCA | DAO Decimal (type 20), size 16 | Precision/scale and authoritative semantics need verification |
| Validation state | Reports.IsValidated | Boolean | Do not set or interpret as Swift readiness without proof |

This sample has additional named address/name components and Chief-related-looking fields; no rule or compatibility behavior is inferred from their presence alone. Do not export using a populated fixture as a supposedly blank template. Obtain a truly empty target-created database and compare object/schema fingerprints first.

The supplied upgrade-tool guide describes a copy-based v33-to-Patch-2 conversion. The tool database was locked by another process when inspection was attempted; checksum verification and code/schema inspection remain pending. No conversion tool code was executed. The supplied checksum's authenticity has not been independently established.

Next evidence: exact NAVFIT Help/About build, verified synthetic values/IDs in every relevant column, folder relation tests, per-form official block mappings, blank template, untouched and changed round trips, all-field reconciliation, validation and print inspection.

## Follow-up after workstation restart

Observed folder rows: Root ID 1 has Parent 0; TEST_FITREP ID 2, TEST_CHIEF ID 3 and TEST_EVAL ID 4 each have Parent 1. Report 1 has null type and Parent `1`; the FitRep, Chief and Eval records have Parent `a 2`, `a 3` and `a 4`, respectively. This establishes sample-specific associations, not the meaning of the `a` prefix or universal join rules. Do not cast Reports.Parent directly to a numeric folder ID. Root is visible in the supplied NAVFIT GUI screenshot, but the screenshot alone does not prove encoded-parent semantics.

The converter became readable after restart. Current SHA256 is `549BCC923328808C804312E96738C8199E964C22FD9667BEF217AF8CE1B8E5E8`; supplied checksum is `4552E4AF7E59ADCBBEC5E725D59E39A8A9C78E2B32AA7896CAA4FB0B02DE2D70`. They do not match. Opening/saving an Access application may change its hash; no cause is established here. The supplied checksum therefore does not verify this file's current state. Obtain an unmodified source or establish provenance before executing the converter. No converter code was executed by this task.

## Installed version and target-created blank database

The running NAVFIT98A.exe reports file/product version 2.2.0.332. User confirms Help/About version 2.2.0.332 and displayed release date August 9, 2026. This is installation evidence, not a claim about release publication history. Executable path/hash are in `evidence/navfit/installed-executable.json`.

User confirms TEST_DB contains only synthetic information and authorizes test changes. NAVFIT_BLANK.accdb was created in the user's running installation. Read-only inspection preserves its hash (recorded in its schema inventory). Ordered field names/types/sizes match TEST_DB for all three data tables. Folders and Reports each contain one row; Summary contains none. Do not discard these default rows without verifying their meaning in NAVFIT. Full properties/index/default reconciliation remains pending.

The controlled mapping probe changes only Comments in the three typed reports; all 824 cells were checked, exactly three intentional changes found, and the source hash preserved. Actual report block placement and NAVFIT acceptance remain pending. Evidence is beside `build/navfit-proof/NAVFIT_MAPPING_PROBE_TESTED.accdb`. This is development mapping evidence, not a completed import/export round trip.

## Comments block mapping confirmed — 3 October 2026

User opened all three controlled mapping-probe reports in NAVFIT98A 2.2.0.332 without errors and observed the synthetic Comments marker in CHIEFEVAL block 40, E1–E6 EVAL block 43, and FITREP block 41. Thus Reports.Comments placement is confirmed for the three tested reports in this installation. Evidence: `evidence/navfit/comments-mapping-confirmation.json`.

This confirms field placement and report opening only. It does not establish validation, narrative capacity/printed fit, complete bidirectional conversion, or suitability of mapping local DraftNotes automatically. Those remain separate tests and implementation decisions.
