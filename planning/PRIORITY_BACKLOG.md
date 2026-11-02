# Swift NAVFIT priority backlog

Updated 1 November 2026 (0.51.0, schema 36). This lists open work only. Finished
items have been removed; their delivery and verification are recorded in
[PROJECT_STATE_HISTORY.md](PROJECT_STATE_HISTORY.md),
[CURRENT_STATE.md](CURRENT_STATE.md) and Git history. Where part of an item is
done, only the remainder is listed. Product invariants and release gates are in
[CONTRIBUTING.md](../CONTRIBUTING.md).

## Planned features (user decisions, 1 Nov 2026)

### F01 — Hierarchical review chain on one shared draft

**Decision:** a report under review is one shared draft. Each reviewer in the
chain can edit it, and later reviewers have more authority, so their edits count
for more. The chain goes back to an earlier reviewer only when that person pulls
it back or someone sends it back to them, and it then runs upward again from
there. This replaces today's model, where a review accepts one exact saved
revision and any save makes the route stale and forces a restart from step 1
(Codex review W01, W02).

**How it works:**
1. **The holder owns the draft.** While a route is pending, the reviewer at the
   current step is its holder. The holder can save as often as needed, and saves
   no longer make the route stale. Saves by anyone else are refused with "This
   report is with *name* (step 2 of 3). Pull it back to edit." The requester counts
   as step 0. As with review decisions today, this is attribution, not a login.
2. **Forward.** The holder chooses Forward (today's Accept), with an optional
   note. Swift records the revision the holder received and the one they passed
   on. The next reviewer sees what changed in between, using the existing
   revision comparison.
3. **Send back.** The holder chooses an earlier step (by default the previous
   one) and must give a reason. That step becomes the holder, and every step from
   there upward goes back to Pending. Their earlier decisions stay in the history.
4. **Pull back.** A reviewer whose step has already forwarded, or the requester,
   can take the draft back, with a reason. The effect is the same as being sent
   back to that step.
5. **Authority.** When a lower step saves over a field that a higher step last
   changed, the save warns "Overrides a change by *name* (step 3)". The higher
   reviewer sees that override flagged when the draft comes back up. Nothing is
   blocked, but the higher reviewer always knows.
6. **Final step.** The last reviewer's Forward completes the route. Local
   approval then locks the revision current at that moment and records which
   revision it was.
7. **History never claims more than happened.** Each step shows the exact
   revision it forwarded. Swift never says a reviewer accepted changes made after
   their step.

**Group routes:** a route started from a summary group no longer goes stale
when another member's report changes. Its group check covers membership and
archive state only. Group completion keeps the full group stamp, so its RSCA
posting still uses exactly the revisions it shows.

**Individual assignments** (ReviewTask) follow the same rules: the assignee holds
the draft and may save, and finishing records the revision. Existing tasks stay
readable.

**Code touched:**
- `modRouteTemplates`: `RouteIsCurrent`, `DecideRouteStep`, new `SendBackRoute` and `PullBackRoute`, and `ApproveSavedRoute`.
- `modRouting.FinishReview`.
- The central save in `CommitDraft`. Members-grid cell saves in `modUXTraitGrid` skip the holder check but log the change on the report.
- The status display in `UXReportState`, `modUXReadRows` and My Work. It shows "With *name* (step 2 of 3)".
- `modUXReview` and the review page.

**Schema (additive):**
- `RouteRunStep` gains `IncomingRevisionID`, `OutgoingRevisionID` and `PassNumber`.
- `RouteRun` gains `ApprovedRevisionID`.
- A `RouteRunEvent` table records forward, send back, pull back and holder saves, linked to `AuditEvent`.
- When a workspace is imported, a pending route takes its old `RevisionID` as the current step's incoming revision. Finished routes keep their history unchanged.

**Members grid (user decision, 1 Nov 2026):** the summary-group Members grid
sits outside the review chain. Grid edits are always allowed, even while a
report is with a reviewer. Each grid change is logged in that report: who
changed which field and when, with a "changed from the Members grid" source. The
holder sees those changes flagged the same way as an override.

**Acceptance:** native tests of a three-step chain:
- edits by each holder, forwarding, sending back to step 1 and to step 2, a pull back by step 1 and by the requester;
- a refused save by a non-holder, and the override warning;
- final approval locking the right revision;
- a group route surviving an edit to another member;
- import of a schema-36 workspace with pending and finished routes.

Update the user guide's review chapter and REVIEW_ROUTING_GUIDE.

### F02 — Departure tracking and MNA roster import

**Decision:** track departures in two ways: from the PRD, and when a member
disappears from a newer roster. Keep the NRRM RUAD import, add MNA RUAD, and
leave out CFL data.

**Roster snapshots.** Every roster import becomes a dated snapshot. It records:
- the format (NRRM or MNA);
- the file's SHA-256;
- its as-of date (from the file if present, otherwise entered);
- the UICs it covers;
- whether the coordinator confirms it is a complete roster for those UICs.

Only a complete roster can flag absences. A filtered or partial export never
marks the rest of the command as departing.

**Matching every row.** Today every row starts as Skip and the coordinator picks
Create or Refresh per row. Instead, Swift suggests a match for every row, from
the strongest identifier the format carries. The coordinator confirms or
changes each suggestion.
- **NRRM:** the roster name already stored on the member, plus rate. NRRM has no DoD ID or SSN.
- **MNA:** a DoD ID, if its sample shows one.

Matched rows refresh as now, which keeps the per-field review and preservation
of manual values.

**Two departure signals:**
- **PRD.** The roster's PRD is stored per member, as now. A member whose PRD is
  within the lead window (90 days by default, set in Settings) shows as
  "Departure expected *date*".
- **Absent from roster.** A member who was present in the last complete roster
  for a UIC, and is missing from a newer complete roster for the same UIC, becomes
  "Departure pending: confirm date". Nothing is archived or deleted
  automatically.

**Resolving a departure.** The coordinator chooses one of three outcomes:
- **Confirm**, with the actual detachment date. The member stays active until the
  coordinator archives them after their reports are done.
- **Dismiss**, with a reason such as "still aboard, missing from that export".
- **Returned**, when the member reappears in a later roster.

Every choice is kept as history.

**Rank from the roster** (Codex W05, replaces R14). When a refreshed roster's
rate or rank differs from the profile's `RankRate`, the preview shows both. It
offers "Update rank from roster" as a reviewed choice; manual corrections are
kept. Due dates in F03 use the profile rank, so this matters.

**MNA RUAD adapter.** `RUADLoad` detects the format from the header row:
- **NRRM:** keeps its exact 35-column check.
- **MNA:** gets its own exact header check, mapped to the same normalized facts:
  name, rate or rank, paygrade, UIC, PRD and any identifier.

Raw MNA columns are stored with a format tag beside today's NRRM values in
`RUADData`. The profile page shows the columns for that member's format.

The MNA column names are guessed only once a **sanitized MNA sample** exists
(headers and two or three made-up rows). Until it arrives, build the snapshot,
matching and departure parts on NRRM. Add a synthetic MNA fixture to the tests
once the format is known.

**Schema (additive):**
- `RosterImport`: format, hash, as-of date, scope UICs, complete flag, who and when.
- `RosterPresence`: each import's member matches.
- `MemberDeparture`: member, basis (PRD or absent), detecting import, PRD, confirmed date, state, resolver, reason, version.

**Acceptance:**
- complete and partial rosters for one and for two UICs;
- a member absent, then back;
- a PRD inside and outside the window;
- two members with the same name;
- a rank change, with and without a manual correction;
- the NRRM 35-column check unchanged;
- the synthetic NRRM roster the import test already generates, extended for these cases.

### F03 — Reports due on the dashboard

**Decision:** the home screen lists members with upcoming reports. A member
drops off once the right report exists, even as a draft.

**What's due.** This is calculated on refresh, like counseling due work. Nothing
is stored except dismissals.
- **Periodic:** the next date from `PeriodicExpectedDate` for the member's worn
  paygrade, using the same versioned calendar and exemptions as validation. E-9
  and CWO5 have no Regular periodic report after 3 June 2026 (NAVADMIN 132/26);
  reuse the rule counseling already applies.
- **Detachment:** due at the confirmed detachment date, or at the PRD while it
  isn't confirmed (from F02). Swift doesn't decide whether a detachment report is
  required (for example, when a periodic report falls close by). It shows both,
  and the coordinator dismisses the one that isn't needed.

**When it shows.** A due report appears once its date is inside the lead
window: 90 days by default, adjustable from 60 to 120 for the 2–3 month
collection period. It stays until satisfied or dismissed, and shows "Overdue"
after the date.

**When it's satisfied.** It is satisfied by a report that:
- is not archived;
- is linked to the member (`MemberReport`);
- has the matching flag (Regular for periodic, Detachment of Individual for detachment);
- has its period end on the due date. For a detachment, the period end falls inside the window around the detachment date.

A draft is enough. Concurrent-only reports don't count.

**Dismiss.** "Not required", with a reason, is stored in a new `ReportDueDismissal`
table (member, basis, due date, reason, who and when). Undo is available.

**On the home screen** (`frmHome`, `UXDashboardRefresh`):
- A "Reports due (N)" count joins the report, group and work counts.
- A list shows member, rate, kind, due date and status (Overdue, Due in N days).
- Actions: Open member, Create draft and Dismiss.
- A "Departures to confirm (N)" link opens the F02 list.
- A line notes "N members have no rank/rate", because no due date can be worked out for them.

**Create draft** starts a report that is already linked to the member. It sets:
- the period end to the due date;
- the Regular or Detachment flag;
- the period start to the day after the member's last report, when that is known.

For several members at once, the selected rows go to the existing directory
batch creation. That creation already checks the contributor and previews before
applying.

**Acceptance:**
- each paygrade's periodic date across the 2026/2027 calendar change;
- E-9 and CWO5 exemptions;
- a draft clearing the item, and an archived or Concurrent-only draft not clearing it;
- a detachment from a PRD and from a confirmed date;
- dismissal and undo;
- a member without rank;
- home-screen refresh time on the 4,700-report synthetic workspace.

Build order: F02 (snapshots, matching, departures) before F03, which needs its
departure dates. F01 is independent and can go in parallel.

## Codex review of 0.51.0 (1 Nov 2026): bounded fixes

An outside review on synthetic copies. Causes were confirmed in source during
triage. R01–R08 were fixed on 1 Nov 2026 (see
[PROJECT_STATE_HISTORY](PROJECT_STATE_HISTORY.md#codex-review-fixes-r01r08));
the rest are open. Review IDs are in brackets. Each fix needs a
native test that fails before the fix.

| ID | Priority | Work |
| --- | --- | --- |
| R09 | Medium | **A Concurrent-only report hides a counseling suggestion** [D07]. `CounselingMatch` (`modCounseling`) skips a session already entered on any other report, including a Concurrent-only one. `CounselingSuggestion` then offers NOT REQ with "No counseling is recorded". Ignore Concurrent-only reports when matching, and never turn a session that is already in use into NOT REQ. |
| R10 | Medium | **Signing succeeds, but the dialog reports failure when draft disposal fails** [D06]. `UXSignoffSave` commits the signoff before the optional disposal. Report the two outcomes separately and offer to retry the disposal alone. |
| R11 | Medium | **Backup ignores unsaved Members-grid cells** [G06]. Warn about pending cells before backing up, or refuse until they are saved or discarded. |
| R12 | Medium | **Restore doesn't check the backup's own receipt** [G07]. When a `.verification.json` is present, compare its SHA-256 before restoring and say whether the file has changed since backup. |
| R13 | Medium | **My Work reads all review history to show active work** [O01]. Preparing one active item took 5.5 ms with no history and 287 ms with 10,000 old tasks and routes. Limit the active view to candidate rows. Pair this with F01, which changes the same readers. |
| R15 | Low | **Small maintenance gaps:** there is no way to view or restore archived counseling [G08], to retire or replace a review route template [G09], or to re-apply a PDF to the right draft after undoing an import into the wrong one [W04]. |

(R14, rank from the roster, moved into F02. T01 is done: the RUAD import test
generates a synthetic NRRM roster when no sample file is given.)

**Small features from the review:**
- **P01, search:** search that ignores name order and punctuation, a period from/to filter, and saved named filter sets.
- **P02, placeholder check:** an Information finding for Swift's own unfinished starter markers, such as `[MEASURABLE RESULT]`. It never blocks a save or output.
- **P03, snippet target:** insert a snippet into the field being edited, at the cursor.
- **P06, people picker:** pick the rater, senior rater and counselor from the directory.

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
| H01 | **Official-template PDF output: acceptance remaining.** The PDF output, summary letter, group packet and Chief comment corrections are built. Three acceptance checks remain: physical printed fit against the official forms, generation on a workstation without Acrobat, and acceptance of the six-point fallback for the official trait-34 grade-4/5 wording (6.5 point). |
| H03 | **Attended NAVFIT98A print check for the narrative line meter (Matt, attended).** The editor's line meter (a72eed8) reproduces NAVFIT98A 2.2.0.332's print wrap from its code, but no NAVFIT98A print has confirmed it. Matt prints one EVAL and one CHIEFEVAL from NAVFIT98A to PDF. Comments, block 28, block 29 and one CPO performance comment are each filled past their limits (Swift's export, default comments pitch). Check three places where NAVFIT98A's warnings and print boxes disagree: the CPO comments box holds 7 lines but NAVFIT98A warns past 6; block 28's third line is slightly clipped (2.9 lines); a CPO performance comment's second line is slightly clipped (1.98 lines). Confirm that wrapping matches the meter. Decide whether the 12-point `CommentsPitch` layout (EVAL 21 lines, FITREP 18) needs modelling; Swift doesn't store `CommentsPitch` today. Adjust `NAVFITFieldFit` (`modPDFOutput`) and its reference cases from the result. |
| M07 | **Accessibility acceptance.** Keyboard order is confirmed complete. Target the DoD Section 508 minimum (36 CFR 1194 / WCAG 2.0 AA) with a conformance report. The remaining checks are screen readers (NVDA, JAWS), focus and announcements, contrast and high contrast, and display scaling and DPI, covering startup, dialogs, report tools, grids, findings and output. Automated geometry and label checks don't establish accessibility. This waits until the user can take part. |
| M01 | **Signed production distribution.** An unsigned `.accdb` opens on target workstations; signing is wanted for polish. Obtain a DoD PKI code-signing certificate through the command ISSM, sign the VBA project and verify a launch under normal security policy. Then produce and verify a signed `.accde`, and rehearse the signing steps (threat model section 5). Blank and Demo `.accdb` files keep being published. Access Runtime support stays out of scope. |

## Performance and storage

| ID | Work |
| --- | --- |
| M03 | **Remaining performance (revisit later, user decision).** The editor takes about 1.1 s to open, flat with workspace size and mostly page layout. Build-time layout baking was declined for now. A 250-member group PDF packet takes about 46 s, most of it in `OutputTo`. |
| L03 | **Workspace size headroom, step 3: sparse revision values.** Steps 1 (size warning) and 2 (no empty-to-empty change rows) are done. This step adds the column `ReportRevision.ValueEncoding`: no value means a full copy, `1` means a missing field is empty. About 12 `src/` readers move to one accessor. A per-revision field count and snapshot hash replace the detection of missing rows, which also delivers threat-model item 6. Import copies rows unchanged, and old revisions are not converted. The measured saving is 24% on sparse reports, probably 10–20% on real ones. Test every persistence runner, schemas 1–36, mixed revisions, disposal, restore and compare. Not critical: about 10 years of headroom at 1,300 reports a year. |

## Code health

| ID | Work |
| --- | --- |
| M05 | **Names built at run time.** Literal form and control names are checked before every build. Names built at run time, such as `"uxNav" & i` or `"frmContext" & panel`, are checked only where a runtime view contract lists them. |
| T05 | **Intermittent test failures, not yet diagnosed.** `test-import-engine` failed once with error 3059 ("Operation canceled by user") during its 16-minute old-VBA benchmark. `report-field-persistence` failed once with Jet error 3802. `test-contributor-startup` reported "Blank startup omitted setup" once during the 0.51.0 release check. Each passed on retry. Capture diagnostics the next time one recurs. |

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
- Every release imports every earlier schema, in either mode.
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
- **F01's hierarchical review chain** gives each report one holder at a time,
  which reduces concurrent edits to the same report. Design shared-mode
  conflicts around it.
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
- Every earlier schema still imports.
- Native tests cover the provider layer in both modes; shared-mode tests run
  against a synthetic test site only.

See [roadmap](ROADMAP.md) for the overview.
