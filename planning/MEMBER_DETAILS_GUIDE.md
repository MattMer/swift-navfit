# Member identifiers, component selections and date reported

> Historical service/verification reference. Current procedures are in the
> [user guide](../docs/Swift_NAVFIT_User_Guide.md); current architecture and
> decisions are in [PROJECT_STATE.md](../PROJECT_STATE.md). Retained field/source
> details remain evidence, but old screens and test commands may be superseded.

Swift NAVFIT 0.16.0 adds three features in a Member details panel for all three report types.

## Member identifiers

SSN supports up to nine digits and DoD ID up to ten, matching the inspected NAVFIT Reports.SSN and DoDID text capacities. Leading zeros are retained. Partial identifiers and blanks remain saveable. Non-digit, formatted or oversized entries reject the entire save without conversion or truncation. These are storage/character checks, not proof of identifier validity. The fields are separate from the local member display name and are not added to default library columns.

## Component selections

Active, TAR, Inactive and AT/ADSW each accept Yes, No or blank. Labels correspond to the supplied schema's Active, TAR, Inactive and ATADSW Boolean fields; TAR is retained as the reference-field label. No combination is inferred or restricted in this build. Swift stores explicit text choices to preserve unentered versus No; external Boolean mapping and component eligibility remain pending.

## Date reported

Enter a real date as yyyy-mm-dd, or leave it blank. Leap dates are accepted; impossible or ambiguous dates and timestamps reject the save. Swift stores a canonical date-only text value, separate from report-period dates, with no automatic default. The inspected source DateReported field is DateTime; export mapping and date-policy/continuity checks remain pending.

## Editing and history

Open Member details, edit and Apply pending values and return. Apply transfers all seven panel values and creates a recovery checkpoint; Save draft commits them with the report. Cancel preserves pending editor values. Unapplied panel edits are not checkpointed. Raw invalid pending identifiers, flags and dates survive recovery without being committed. Archived drafts can be viewed but cannot apply edits.

All seven fields participate in snapshots, comparison, duplication, restore and stale-review checks. Old snapshots display Not recorded; restoring them clears unrecorded fields to null. New schema-15 snapshots require all seven rows, including explicit nulls. Checklist prompts assess presence only; incomplete saves remain available. No identity authenticity, component eligibility or Navy validation claim is made. NAVFIT import remains narrative-only and full mapping/export remains pending. Printed layout is deferred.

Use scripts/upgrade-workspace.ps1 with the new blank template and a fresh destination for existing populated workspaces. Verification: 125 new assertions plus 881 prior behavioral assertions, 169 catalogue checks and all 2472 retained import values passed. Schema 14-to-15, 15-to-15 and 5-to-15 migrations compare all original rows/fields, preserve source hashes and reopen. All 20 distribution user-data tables are empty. Runtime/x86 and comprehensive visual verification remain pending.
