# r115 â€” workspace data transfer and fixed report pages

App 0.25.0-dev; schema 25. Existing workspaces and earlier distributions are preserved.

## Transfer data into a newer release

Open the new **blank** release. Choose **Import data from a previous Swift NAVFIT version** on first-run setup, or under **Settings & Help** after initial setup. Select the old Swift `.accdb`, inspect the source/table counts, and choose **Import all workspace data**. Close and reopen the new app after success.

This is a full workspace transfer, not a report-only import or a merge of independent workspaces. It preserves IDs, relationships, report values/revisions, local reviews/routes/approvals, notes, recovery, defaults, directory/prior evaluations, reporting seniors/RSCA history, imports/provenance and operation history. Initial contributor/workspace setup in an otherwise blank destination is replaced by the source setup. A destination containing other user data is rejected rather than overwritten or silently merged by ID. The synthetic demo is therefore not a migration target.

The source is opened through a separate read-only DAO workspace. No source forms, VBA or macros are executed or copied. Production migration verifies every original row/field before commit; an outer transaction restores initial destination setup if transfer fails. Source hashes are checked before and after transfer. Unsupported future schemas, unknown source tables and changed source files stop the operation. This engine is also used by `scripts/upgrade-workspace.ps1`.

Future release requirements are recorded in `AGENTS.md`: additive schema changes, retained historical fixtures, exact transfer checks and rejection/rollback gates before publication.

## Report workspace

- **Front Page** and **Back Page** replace the seven official section tabs. Draft notes, Review and History remain separate local pages. Legacy detail/validation links select the corresponding physical page.
- Fixed control geometry follows official block order with form-specific page boundaries. The supplied 2025 NAVPERS XFA page subforms place EVAL/FITREP traits 38â€“39 on the back and CHIEFEVAL traits 37â€“39 on the back. All nonlegacy official fields remain available; structured export fields remain in the existing hidden buffer.
- Window changes adjust the viewport and action-bar position, not report field sizes or column layout. A narrower window can scroll the fixed canvas. Comments retain their existing fixed Courier New/plain-text dimensions; this does not establish printed fit.
- Page changes flush active text into the existing unbound buffer without creating a saved revision. Save, automatic recovery, undo, approval locks and error navigation retain their services.
- Opening no longer performs the old section layout repeatedly. Initial complete live validation is deferred to the first dirty timer tick; field-exit and explicit validation/save checks remain available.

## Tables and groups

Summary groups show **Closeout date**, meaning report period ending date. Matching member ending dates show one ISO date; missing dates show **Incomplete**, and differing dates show **Mixed**. An empty group can show its effective ending-date default.

Multi-column list tables expose native or added clickable headings. The first click sorts ascending and a repeat reverses direction. Text, numeric/date values, filters and bound selections are retained. Value-list heading rows and quoted semicolons are preserved. Changing the list source resets the next click to ascending. Editable member-grid sorting drains pending cells before changing row order and keeps report IDs; sorting itself creates no report revisions. The RSCA entry table reorders controls while retaining pending values and keyboard order.

## Verification and remaining acceptance

Native scratch-copy checks cover page/field availability for all three types, pending-buffer preservation, fixed geometry, local workflow/error links and closeout behavior; sorting, save/recovery timers and the full import engine are checked separately. Retained actual historical distributions are used, rather than editing metadata to simulate old schemas. Most earliest distributions are blank; populated older and current fixtures exercise actual data/history copying.

Controlled four-open benchmark, with the editor verified closed between each measurement: r114 median 2.59 seconds; candidate median 1.13 seconds. The deferred initial validation took 38â€“51 milliseconds. This measures one synthetic report on this computer, not an end-user or large-group guarantee.

Manual fixed-page visual inspection, actual header mouse interactions across displays, keyboard/screen-reader behavior and scaling remain acceptance work. No computer-use automation or verified printing claim is made.

546 relevant native assertions/binding checks passed: 240 fixed-page/buffer/closeout, 38 sorting, 53 persisted heading bindings, 28 efficiency, 119 resource, 30 help-consistency and 38 combined member/PSG/RUAD checks. The real 30-second recovery checkpoint passed separately. The 25-schema migration matrix, populated schema-20/history and same-schema migrations passed, including source hashes and rollback/rejection guards. Early schema distributions are mostly blank; populated historical fixtures are identified in the manifest.

Both final builds compile/start; the blank has 48 checked user-data tables empty and the demo contains nine synthetic reports. [Verification manifest](WORKSPACE_COMPAT_PAGES_R115_VERIFICATION.json) records exact local SHA256, byte sizes, source hashes, transfer fixtures and benchmark samples.

- [Blank r115](https://drive.google.com/file/d/1lWyVjr-oAcKE59fQJZ24JZYak37o9CfE/view?usp=drivesdk).
- [Synthetic demo r115](https://drive.google.com/file/d/1UNpW5NX8DxsgesiZ2C7_mSaq3uEsZgQN/view?usp=drivesdk).

Drive metadata readback verified exact names, byte sizes and the established parent folder. No remote checksum was returned; sharing was not changed. Earlier files and user-edited workspaces remain untouched. GitHub was not pushed in this change.
