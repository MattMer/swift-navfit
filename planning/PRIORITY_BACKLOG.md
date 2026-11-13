# Swift NAVFIT priority backlog

Updated 10 November 2026 (0.55.0 released; 0.56.0-dev, schema 48). Manual check
results from 10 November are folded in below. This lists open work only. Finished
items have been removed; their delivery and verification are recorded in
[PROJECT_STATE_HISTORY.md](PROJECT_STATE_HISTORY.md),
[CURRENT_STATE.md](CURRENT_STATE.md) and Git history. Where part of an item is
done, only the remainder is listed. Product invariants and release gates are in
[CONTRIBUTING.md](../CONTRIBUTING.md).

## EVAL block 44 (9 Nov 2026)

**Released in 0.55.0** (schema 46 onward, field catalogue 3): EVAL
block 44 (Qualifications/Achievements) on the Back Page, NAVFIT98A import/export
(`Qualifications`), PDF import (XFA `BB`), PDF output with an added page for
overflow, and restore from retained NAVFIT98A imports when an earlier workspace
is imported. **Still open:** NAVFIT98A's own block 44 print box and limit are
not measured (no line meter, only Swift's PDF fit finding); confirm the printed
fit of `BB` on paper.

## Planned features (user decisions, 1 Nov 2026)

### F01 — Hierarchical review chain on one shared draft

**Implemented in development** (0.52.0-dev, schema 37, 2 Nov 2026); see
[PROJECT_STATE_HISTORY](PROJECT_STATE_HISTORY.md#f01-shared-draft-review-chain).
Shipped in 0.52.0. Remaining: route events are not linked to `AuditEvent`.
The user guide PDF is regenerated at release.

### F02 — Departure tracking and MNA roster import

**NRRM and MNA implemented in development** (0.52.0-dev, schema 38, 2 Nov 2026); see
[PROJECT_STATE_HISTORY](PROJECT_STATE_HISTORY.md#f02-roster-snapshots-and-departures).
Shipped in 0.52.0.

**Still open:**
- **Officer rank from MNA.** MNA lists the designator but not the rank, so a
  new officer profile without a linked report starts without one. The current
  source already fills a missing rank from the latest linked report.
- **Real PRD format (checked 10 Nov, DATA-01).** A real NRRM export shows PRD as
  a date cell displayed dd-Mon-yy (for example 31-Dec-30); as a plain number the
  same cell is the spreadsheet date serial (47848). `RosterPRDDate`
  (`modRoster`) reads neither: add dd-Mon-yy and dd-Mon-yyyy, the date serial,
  and a value that already arrives as a date.

### F03 — Reports due on the dashboard

**Implemented in development** (0.52.0-dev, schema 39, 2 Nov 2026); see
[PROJECT_STATE_HISTORY](PROJECT_STATE_HISTORY.md#f03-reports-due).
Shipped in 0.52.0. The lead window stays fixed at 90 days by owner decision.
Remaining: Create draft makes one report at a time rather than sending several
to directory batch creation.



## Review decisions not planned

R01-R13 were completed earlier. R15, P01 and M05 shipped in 0.53.0 (schema 44).
R14's remaining input-data limitation is recorded under F02.

**Not taken:**
- **P04, accomplishment notebook:** waits for demand.
- **P05, "Why this value?":** low value for its cost.
- **P07, restoring selected fields from history:** low value for its cost.
- **CFL PRT import:** the user decided against it, 1 Nov 2026.
- **The rest of the review's input-package tracker [G04]:** depends on the CFL import, so it is not planned.
- **Regular versus Concurrent on prior evaluations [G05]:** a known limitation, not planned.

## Output and acceptance

| ID | Work |
| --- | --- |
| H01 | **Official-template PDF output: acceptance remaining.** Passed on paper (10 Nov): EVAL, FITREP, the summary letter and group packet, and the 6.5-point trait-34 grade-4/5 wording (accepted). PDF output works on a Navy-managed workstation after unblocking the file and enabling content (PRT-08). Remaining: (1) **CHIEFEVAL fixes (PRT-03 failed):** block 51 lines print too far apart, and the block 20 physical readiness result does not print next to its MM/YY date. (2) PDF creation on a workstation without Acrobat (PRT-07). (3) Recheck on the Navy workstation with a physical default printer: 0.55.0 split pages there; 0.56.0-dev (1af7fd6) lays out on Microsoft Print to PDF. |
| H03 | **Attended NAVFIT98A print check for the narrative line meter (Matt, attended).** The editor's line meter (a72eed8) reproduces NAVFIT98A 2.2.0.332's print wrap from its code, but no NAVFIT98A print has confirmed it. Matt prints one EVAL and one CHIEFEVAL from NAVFIT98A to PDF. Comments, block 28, block 29 and one CPO performance comment are each filled past their limits (Swift's export, default comments pitch). Check three places where NAVFIT98A's warnings and print boxes disagree: the CPO comments box holds 7 lines but NAVFIT98A warns past 6; block 28's third line is slightly clipped (2.9 lines); a CPO performance comment's second line is slightly clipped (1.98 lines). Confirm that wrapping matches the meter. Decide whether the 12-point `CommentsPitch` layout (EVAL 21 lines, FITREP 18) needs modelling; Swift doesn't store `CommentsPitch` today. Adjust `NAVFITFieldFit` (`modPDFOutput`) and its reference cases from the result. |
| M07 | **Accessibility acceptance.** Keyboard order is confirmed complete. Target the DoD Section 508 minimum (36 CFR 1194 / WCAG 2.0 AA) with a conformance report. Passed 10 Nov: Windows high contrast on all forms, menus and pop-ups (A11Y-06), and display scaling at 125% and 150% (A11Y-07). Remaining: screen readers (NVDA, JAWS) with focus and announcements across startup, dialogs, report tools, grids, findings and output, and a colour contrast spot check. Automated geometry and label checks don't establish accessibility. This waits until the user can take part. |
| M01 | **Signed production distribution.** An unsigned `.accdb` opens on target workstations; signing is wanted for polish. Obtain a DoD PKI code-signing certificate through the command ISSM, sign the VBA project and verify a launch under normal security policy. Then produce and verify a signed `.accde`, and rehearse the signing steps (threat model section 5). Blank and Demo `.accdb` files keep being published. Access Runtime support stays out of scope. |

## Screen changes from the 10 Nov review

Pop-up style, list columns, the Reports filters, high contrast and scaling all
passed. The owner asked for these changes:

| ID | Work |
| --- | --- |
| U01 | **Taller Reports list.** Reports shows about 6 rows; make it show 10 (UI-03). |
| U02 | **One name format: LAST, FIRST.** Overview's Reports due shows "Reese Parker" while Recent reports shows "PARKER, REESE". Use "PARKER, REESE" in both, and in the command member directory table (UI-05). |
| U03 | **Contributors moves to Resources.** Remove the Contributors tab from My Work and open it from Resources instead (UI-06). |
| U04 | **Comments box shows every printed line.** Block 43 (EVAL comments) printed inside its box with 18 lines near the limit, but the editor showed 17 before scrolling (PRT-06, which named block 44 by mistake). Fixed in development: each narrative box is now sized from the PDF line count and the screen line height. |
| U05 | **Library's second dark-blue button (decision pending, UI-04).** Decide whether Import reports stays a main action. |

## Performance and storage

| ID | Work |
| --- | --- |
| M03 | **Remaining performance (revisit later, user decision).** The editor takes about 1.1 s to open, flat with workspace size and mostly page layout. Build-time layout baking was declined for now. A 250-member group PDF packet takes about 46 s, most of it in `OutputTo`. |
| L03 | **Workspace size headroom, step 3: sparse revision values.** Steps 1 (size warning) and 2 (no empty-to-empty change rows) are done. This step adds the column `ReportRevision.ValueEncoding`: no value means a full copy, `1` means a missing field is empty. About 12 `src/` readers move to one accessor. A per-revision field count and snapshot hash replace the detection of missing rows, which also delivers threat-model item 6. Import copies rows unchanged, and old revisions are not converted. The measured saving is 24% on sparse reports, probably 10–20% on real ones. Test every persistence runner, schemas 1–36, mixed revisions, disposal, restore and compare. Not critical: about 10 years of headroom at 1,300 reports a year. |

## Code health

| ID | Work |
| --- | --- |
| T05 | **Intermittent test failures, not yet diagnosed.** `test-import-engine` failed once with error 3059 ("Operation canceled by user") during its 16-minute old-VBA benchmark. `report-field-persistence` failed once with Jet error 3802. `test-ux-admin-feedback` failed twice in the 2 Nov extended run ("Reviewed proposal apply failed" with an empty error) and passed on three re-runs. `test-contributor-startup` reported "Blank startup omitted setup" once during the 0.51.0 release check. On 3 Nov 2026 it failed once with "Reserved Error" from `Eval` after all 61 assertions had passed, and passed on two reruns. Each passed on retry. Capture diagnostics the next time one recurs. On 12 Nov 2026 native runs moved to a hidden Windows desktop because the failures clustered while the PC was in use; this is an unconfirmed hypothesis, and T05 closes after three consecutive ExtendedTest runs on a PC in use report no FLAKY entry. |
| T06 | **Remove pre-35 import code with the next schema change (owner decision, 10 Nov 2026).** The importer now accepts schema 35 onward. In `modMigration`, drop the copy/check skip lines for schemas below 35, the always-true `targetVersion >= 22..34` checks, the `ver >= 26/28/31` wrappers and the pre-28 Report column branch (about 60–70 lines). Retire `test-persistence-transfer` and `test-audit-message-transfer` (built for schemas 21–33) and the 0.24.0-preview default template in `upgrade-workspace.ps1`. Keep `MigrateLegacyReportNotes` and `MigrateDirectoryCompatibility` (Demo build and startup use them), the legacy columns and the schema 1–34 definition files. Verify with the full release gate that the schema change needs anyway. |

## M11 — Optional SharePoint Lists workspace storage (planning only)

**Status:** proposed and not scheduled. Today Swift NAVFIT is single-file and
offline: one authoritative workspace `.accdb`, one contributor at a time
([PROJECT_STATE.md](../PROJECT_STATE.md)), and no network interfaces
([SV-1](../docs/architecture/views/SV-1.md)). Those statements stay true until
this work ships. No other backlog item covers shared or multi-user storage.

**Objective:** each command chooses a storage mode.
- **Local mode (default):** reports and supporting data remain inside the Swift
  NAVFIT `.accdb`, exactly as today.
- **Shared mode:** each user runs a local Access front end connected to the
  command's SharePoint Lists.

The report editor, validation, summary groups, rack and stack, comments,
revisions, routing, and NAVFIT98A import/export remain available in both modes.

**Requirements:**

1. **Setup/settings screen.** A workspace setup and settings screen shows the
   active command, the workspace, and the storage mode.
2. **Discovery.** For SharePoint mode, the user enters a site URL and Swift
   discovers an existing Swift NAVFIT workspace on that site. Users never choose
   individual lists.
3. **Provisioning.** If the workspace doesn't exist, Swift offers guided
   provisioning of the required lists, columns, indexes, configuration records
   and supported settings, subject to the user's SharePoint permissions.
4. **Detection and recovery.** Swift detects the schema version, incompatible
   structures, missing components and insufficient permissions. Interrupted
   provisioning can be resumed, and never overwrites existing command data.

**Product invariants that still apply** ([CONTRIBUTING.md](../CONTRIBUTING.md#product-invariants-non-negotiable)):
- Every release imports every supported schema, in either mode.
- Schema changes are additive and versioned in both modes, and the SharePoint
  schema carries its own version, checked against the front end's.
- Report history is never deleted, except confirmed draft disposal.
- A populated workspace is never silently replaced, and records are never merged
  by ID. Moving between modes uses the production import into an empty target.
- The fixed report-editor geometry, the unbound buffer and recovery are unchanged.
  Pending-edit recovery checkpoints stay local to each user's front end.
- Demo and test data stay synthetic. Never put personnel data on a test site.

**Constraints of SharePoint Lists that the design must address:**

| Area | Today (local `.accdb`) | With SharePoint Lists |
| --- | --- | --- |
| Transactions | A save, its revision, changes and receipt commit together | No transaction across lists or items. Needs an idempotent write protocol (a commit marker or journal per save) and reconciliation after interruption. |
| Integrity rules | Table validation rules, required fields, unique indexes, enforced relationships | Not enforced the same way: limited unique columns, lookup relationships, no table-level rules. Swift must enforce them in its data layer and verify on read. |
| Identity | AutoNumber IDs are referenced throughout | SharePoint assigns its own item IDs. Use Swift's existing GUIDs as stable keys and map IDs. |
| Size and indexing | No practical row limit for a binder | List view threshold (5,000 items) and per-list index limits. Queries must stay on indexed columns, with archiving guidance. |
| Concurrency | One contributor at a time | Several users at once. Optimistic concurrency on every write (extending the existing version counters, or using item versions) with clear conflict messages. |
| Attribution | Contributor choice; attribution only | SharePoint records an authenticated "Modified By". Decide whether the contributor follows the signed-in account. |
| Performance | Local disk | Network latency. Batched reads and caching matter more. Define offline behavior. |
| Security | Folder permissions on approved storage | SharePoint permissions, plus authorization of the tenant and site for PII/CUI. Decide whether SSNs are stored there at all. |

**Decisions (product owner, 26 Oct 2026):**

| Topic | Decision |
| --- | --- |
| Environment | DoD365 SharePoint Online, the command's site |
| SSNs | Stored in SharePoint if the ISSM approves; otherwise shared mode keeps them out |
| Offline | Read-only cache: users can view cached data but not save until reconnected |
| Attribution | In shared mode, the contributor is the signed-in account; local mode keeps the startup profile choice |
| Moving between modes | Both directions (local to shared, shared to local), each through the verified import into an empty target |
| Provisioning | From inside Swift's setup screen, using the method the research spike proves reliable on DoD365 |
| Who may set up or migrate | SharePoint site owners (list-management rights); Swift checks those permissions first |

**Still open (needs the ISSM):**
- Confirmation that the command's DoD365 site is authorized for this PII/CUI.
- Whether SSNs may be stored there. Until approved, plan for the no-SSN variant.

**Design notes for planning:**
- A **storage-provider layer** behind one interface, so the editor, validation,
  groups, routing, comments and NAVFIT98A exchange don't change with the mode.
  Local mode keeps the current tables. Shared mode uses linked SharePoint lists.
- **A configuration list** marks a Swift NAVFIT workspace. It holds the workspace
  GUID, command, schema version, minimum front-end version and a provisioning
  journal. Discovery looks for it, so users never pick lists.
- **Provisioning is resumable.** Create components in dependency order and record
  each completed step in the journal. A rerun verifies what exists (by internal
  name and Swift marker), creates only what's missing, and stops with a diagnosis
  on any incompatible structure. It never drops, renames or rewrites a list that
  holds data.
- **Pre-checks before any write:** permissions (provisioning needs list
  management rights; daily use needs contribute rights), schema version and
  front-end compatibility.
- **F01's review chain** gives each report one holder at a time, which reduces
  concurrent edits to the same report. Design shared-mode conflicts around it.
- **Constraints from the 1 Nov 2026 Codex review (A01–A08), for the phase-1 spike:**
  - SharePoint batch requests are not all-or-nothing, so a save needs a verified
    publish step, and an ETag compare alone doesn't make a write atomic.
  - `Workspace.ActiveProfileID` is a single shared row, so contributor sessions
    must stay local.
  - Linked lists allow offline edits by default, so the read-only offline
    policy has to be enforced.
  - `dbOpenTable`/`Seek` revision readers fail on linked tables.
  - 14 `AuditEvent` relationships and the row-based revision history need
    deliberate list and query mapping.
  - The retired ACS sign-in patterns are unusable.
  - Today's backup rejects linked tables, so shared mode needs its own backup
    and restore contract.

**Testing strategy (product owner, 26 Oct 2026):** SharePoint has no container
image and no open-source implementation of its list protocol, so testing uses
three layers:

1. **Simulated SharePoint provider, every build, no cost.** A test-only provider
   on local Access tables that behaves like SharePoint Lists:
   - no transactions across lists;
   - relationships, validation rules and most unique constraints not enforced;
   - its own item IDs;
   - the 5,000-item view threshold and per-list index limits;
   - optional throttling;
   - injected failures: provisioning cut off at step N, a write dropped midway,
     a concurrent edit by another user.

   The same suite runs against the Local and Simulated providers in the extended
   tier. It covers idempotent saves, recovery, resumable provisioning,
   version and compatibility checks, conflicts and integrity checks. It can't
   cover Access's SharePoint connector or sign-in.
2. **A small commercial Microsoft 365 tenant for the connector.** Business Basic
   (trial or monthly), which runs the same SharePoint Online code as DoD365, apart
   from sovereign-cloud endpoints and the command's sign-in policies. It hosts the
   phase-1 research spike and a manual-tier runner against a dedicated test site.
   Synthetic data only.
3. **A DoD365 sandbox site for final acceptance.** A separate test site or
   subsite on the command's DoD365 with synthetic data only, requested from the
   site owner or ISSM. It is the only place to prove DoD365 sign-in and
   conditional-access policies.

Not chosen: a local SharePoint Server Subscription Edition lab. Its 180-day trial
needs Windows Server, SQL Server and a domain controller (16 GB+ RAM), and it
differs from SharePoint Online in sign-in and limits.

**Phasing:**
1. A research spike on a synthetic test site: linked-list behavior in Access,
   write atomicity and recovery, list limits, the provisioning API, sign-in.
   Record the findings and decisions here.
2. The storage-provider layer in local mode only, with no behavior change, plus
   the simulated SharePoint provider (testing layer 1). The full extended tier
   must pass unchanged.
3. The setup/settings screen (requirement 1), showing local mode.
4. SharePoint discovery, version and compatibility checks, and permission checks
   (requirements 2 and 4).
5. Guided, resumable provisioning (requirement 3), including interruption tests.
6. Shared-mode reads and writes with concurrency and conflict handling.
7. Moving data between modes through the production import, verified value by
   value.
8. Acceptance on the approved environment. Update the threat model, SV-1, SBOM,
   DevSecOps matrix, user guide, README and PROJECT_STATE.

**Acceptance criteria:**
- Local mode behaves exactly as today.
- Every listed feature works in both modes.
- Provisioning interrupted at any step resumes without data loss.
- An incompatible or newer workspace is refused with a clear diagnosis.
- Insufficient permissions are detected before any change is made.
- Two concurrent users can't silently overwrite each other's edits.
- Every supported schema still imports.
- Native tests cover the provider layer in both modes; shared-mode tests run
  against a synthetic test site only.

See [roadmap](ROADMAP.md) for the overview.
