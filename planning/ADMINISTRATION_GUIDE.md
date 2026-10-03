# Administrative draft fields

Swift NAVFIT 0.9.0 adds a first administrative slice to all three draft types. Near the top of the editor, below the local display name, enter:

| Swift field | Capacity | Inspected NAVFIT Reports field |
| --- | ---: | --- |
| Official full name | 27 characters | FullName |
| Grade / rate | 5 characters | Rate |
| Designator / designation | 12 characters | Desig |
| Unit UIC | 5 characters | UIC |
| Ship / station | 18 characters | ShipStation |

These capacities are measured from the supplied NAVFIT data schema and match the bundled empty template. They are storage constraints, not proof of official format, permissible codes or NAVFIT round-trip compatibility. No automatic mapping/import/export is added. Administrative fields are manually entered, single-line text or null. Tabs and line breaks are rejected. Leading zeros, punctuation and case remain intact; overlength values reject the save without truncation. Incomplete and blank drafts remain saveable.

The local display name stays independent of the official full name. It is not converted, copied or reformatted automatically. SSN, DoD ID, component/name parts, report occasions, duty/observation categories, reporting-senior administration and traits remain pending.

Save draft commits these fields with the existing report fields in one revision. Case-only changes are recorded; no-op saves do not create history. Duplication copies administration into the new report. Restore creates a new revision using the selected snapshot. Restoring a revision that predates administration clears these fields to null rather than inventing values. Legacy comparison shows Not recorded; new snapshots carry schema 9 and require all five administrative snapshot rows, including explicit nulls.

Revision comparison includes all 12 current report fields. Administrative saves make previous review assignments stale through their revision reference. Review comments remain separate. Narrative-only reconciliation preserves administration; undo restores the relevant saved snapshot.

Timed/manual editor checkpoints now retain administrative edits too, including overlength entries that cannot yet be saved. Recovery displays the pending values without committing them; shorten invalid values deliberately before saving. Checkpoints from older schemas have null administrative entries. The saved-draft checklist adds five presence prompts, expressly distinct from Navy policy requirements; it excludes unsaved editor values.

Use scripts/upgrade-workspace.ps1 to carry an existing workspace into the new schema-9 app at a fresh path. Keep the supplied blank distribution separate from your populated workspace. Earlier report revisions are retained exactly, without fabricating administrative snapshots.

Verification: 53 administrative assertions cover all three types, leading zeros, punctuation, exact boundaries, oversized/nontext inputs, null/empty values, no-op preservation, duplication/restore, case changes, stale reviews, late snapshot failure rollback, editor save, comparison, legacy history and recovery in a new Access process. All existing narrative, comparison, recovery, workflow, provenance, comments, routing, smoke and direct-review suites pass: 350 assertions total. Narrative import verifies all 2472 retained values and rollback. Schema-8-to-9 and schema-9-to-9 migrations compare every original row/field and reopen; schema-5-to-9 additionally preserves older recovery checkpoints. Full visual/Runtime/x86 and full NAVFIT round-trip verification remain pending.
