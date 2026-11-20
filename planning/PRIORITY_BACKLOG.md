# Swift NAVFIT priority backlog

Updated 19 November 2026 (0.59.0, schema 51). Manual check
results from 10, 14 and 15 November are folded in below. This lists open work only. Finished
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
is imported. Swift's printed fit of block 44 passed on paper (14 Nov). **Still open:**
NAVFIT98A's own block 44 print box and limit are not measured (no line meter,
only Swift's PDF fit finding).

## Review decisions not planned

R01-R13 were completed earlier. R15, P01 and M05 shipped in 0.53.0 (schema 44).
R14's remaining input-data limitation is recorded under F02.

**Not taken:**
- **Linking route events to `AuditEvent` (F01), officer rank from MNA (F02),
  and several drafts at once from Reports due (F03):** not planned.
- **P04, accomplishment notebook:** waits for demand.
- **P05, "Why this value?":** low value for its cost.
- **P07, restoring selected fields from history:** low value for its cost.
- **CFL PRT import:** not planned.
- **The rest of the review's input-package tracker [G04]:** depends on the CFL import, so it is not planned.
- **Regular versus Concurrent on prior evaluations [G05]:** a known limitation, not planned.

## Output and acceptance

| ID | Work |
| --- | --- |
| H01 | **Official-template PDF output: acceptance remaining.** Passed on paper: EVAL, FITREP, the summary letter and group packet, the 6.5-point trait-34 grade-4/5 wording (10 Nov), the summary packet starting reports on a new sheet when printed on both sides, the Navy workstation with a physical default printer (14 Nov), and CHIEFEVAL with block 20 printed once (PRT-03, 15 Nov). Remaining: PDF creation on a workstation without Acrobat (PRT-07). |
| M07 | **Accessibility acceptance.** Keyboard order is confirmed complete. Target the DoD Section 508 minimum (36 CFR 1194 / WCAG 2.0 AA) with a conformance report. Passed 10 Nov: Windows high contrast on all forms, menus and pop-ups (A11Y-06), and display scaling at 125% and 150% (A11Y-07). Colour contrast spot check passed (14 Nov). Remaining: screen readers (NVDA, JAWS) with focus and announcements across startup, dialogs, report tools, grids, findings and output. Automated geometry and label checks don't establish accessibility. This waits until the user can take part. |
| M01 | **Signed production distribution.** An unsigned `.accdb` opens on target workstations; signing is wanted for polish. Obtain a DoD PKI code-signing certificate through the command ISSM, sign the VBA project and verify a launch under normal security policy. Then produce and verify a signed `.accde`, and rehearse the signing steps (threat model section 5). Blank and Demo `.accdb` files keep being published. Access Runtime support stays out of scope. |

## Screens

U01–U04 passed on screen (14 Nov), and the Reports archive filter starting on Active (U06, 15 Nov). The Library keeps both main actions (U05).

## Performance and storage

| ID | Work |
| --- | --- |
| M03 | **Remaining performance (revisit later).** The editor takes about 1.1 s to open, flat with workspace size and mostly page layout. Build-time layout baking was declined for now. A 250-member group PDF packet takes about 46 s, most of it in `OutputTo`. |

## NAVFIT98A exchange

Unanswered check boxes export unticked (X01), the NAVFIT98A narrative print check (H03) and the ISIC summary letter profile in group export (X02) passed (15 Nov) and shipped in 0.57.0. No open items.

## Code health

No open items. T05 (intermittent native test failures) closed on 19 November 2026 after three consecutive release checks with no flaky step: 0.57.0, 0.58.0 and 0.59.0.

## S01 — Front end and data file

**Released in 0.56.0 (schema 49).** The application and the workspace data
are separate files. Setup makes `Swift NAVFIT data.accdb` beside the
application (Change folder picks another place); a data file of that name
beside the application is used without asking. A newer version upgrades an
older data file in place after a checked copy, and puts it back from the copy
if a step fails. A missing or unusable data file opens a window with Find the
data file, Create a new data file and Close Swift NAVFIT. Open an existing data
file brings a complete single-file workspace from 0.49.0 to 0.55.0 into a new
data file beside it. Backup copies the data file and checks it row by row.
Releases ship as `Swift_NAVFIT_<version>.accdb` plus the Demo pair
(`scripts/package-release.ps1`). Each application keeps its own contributor,
saves retry on a lock, presence shows on the Overview and in the editor
(`SessionPresence`), and sync conflict copies are named at startup.

**Several people on one data file (0.60.0-dev).** Up to 10 people at once on a
data file on a mapped network drive are supported; a synced
Teams/SharePoint/OneDrive folder is not (each computer gets its own copy; M11
covers shared storage in SharePoint Lists). A stale or conflicting save is
refused with the entries kept, and lock conflicts retry. Startup lists who else
is connected from the lock-file roster and warns when the chosen contributor is
active on another computer; presence is cleared on exit. Backup retries, then
refuses cleanly, while others save. A coordinator **Settings & Help > Compact
data file** runs only when nobody else is connected. Upgrade refusals name who
is connected, and an older copy on an upgraded file names the version to get.
An automated test runs 8 sessions on one data file.

**Attended checks (on a fresh 0.56.0 copy in an ordinary folder).** Checks 1, 2, 3 and 5
passed (14 Nov). Check 4 (setup's **Restore a backup or import from a previous version**)
and check 8 passed (15 Nov). Open: 6 and 7.
1. First setup asks for the UIC and makes the data file beside the
   application; a report saves.
2. Rename or move the data file and reopen: the window names the missing file;
   Find the data file, Create a new data file and Close each work.
3. Open an existing data file on a copy of a 0.55.0 workspace: the workspace
   comes in and the 0.55.0 file is unchanged.
4. Settings & Help backup on the data file.
5. The Demo pair opens when both files are saved in the same folder.
6. Several computers (3 or more) on one data file on the command's mapped
   drive: all open Swift NAVFIT at once; each works on their own report, then
   two edit the same report (a refused save keeps the entries); two edit a
   group's cells together; a backup is made while others work; after closing,
   the Overview list clears. The README "Not yet tested" network-share item
   stays until this passes.
7. A data file in OneDrive with a forced conflict copy: startup names the copy.
8. The upgrade message, at the next schema change (or on a copy whose
   `AppMetadata.SchemaVersion` is set back to 48).

**Known limits:** a copy that is already linked cannot start a second data file
(a new download can). A session that ends without closing Access (a crash or power
loss) stays listed for up to 3 minutes. A change to a table validation rule set by `build.ps1` will need
upgrade-step support when it first happens.

## M11 — Optional SharePoint Lists workspace storage (planning only)

**Status:** proposed and not scheduled. Today Swift NAVFIT is an application
plus a separate data file, offline: one authoritative data file per command,
used by up to 10 people at once on a mapped network drive, each with their own
application copy and contributor ([PROJECT_STATE.md](../PROJECT_STATE.md)), and
no network interfaces
([SV-1](../docs/architecture/views/SV-1.md)). Those statements stay true until
this work ships. No other backlog item covers shared storage beyond a mapped
network drive.

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

| Area | Today (local or network-share `.accdb`) | With SharePoint Lists |
| --- | --- | --- |
| Transactions | A save, its revision, changes and receipt commit together | No transaction across lists or items. Needs an idempotent write protocol (a commit marker or journal per save) and reconciliation after interruption. |
| Integrity rules | Table validation rules, required fields, unique indexes, enforced relationships | Not enforced the same way: limited unique columns, lookup relationships, no table-level rules. Swift must enforce them in its data layer and verify on read. |
| Identity | AutoNumber IDs are referenced throughout | SharePoint assigns its own item IDs. Use Swift's existing GUIDs as stable keys and map IDs. |
| Size and indexing | No practical row limit for a binder | List view threshold (5,000 items) and per-list index limits. Queries must stay on indexed columns, with archiving guidance. |
| Concurrency | Up to 10 people at once on a mapped network drive; stale saves refused, lock conflicts retried | Several users at once without a shared drive. Optimistic concurrency on every write (extending the existing version counters, or using item versions) with clear conflict messages. |
| Attribution | Contributor choice; attribution only | SharePoint records an authenticated "Modified By". Decide whether the contributor follows the signed-in account. |
| Performance | Local disk | Network latency. Batched reads and caching matter more. Define offline behavior. |
| Security | Folder permissions on approved storage | SharePoint permissions, plus authorization of the tenant and site for PII/CUI. Decide whether SSNs are stored there at all. |

**Decisions:**

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
- **Constraints from the design review (A01–A08), for the phase-1 spike:**
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

**Testing strategy:** SharePoint has no container
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
