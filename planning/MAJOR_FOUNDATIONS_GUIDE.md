# Major foundations checkpoint — 0.21.0

Schema 20 adds eighteen tables for explicit summary groups, reviewed bulk changes, routes, meetings and versioned validation. The draft editor still saves 78 fields. Use Home to open the feature workbench, PDF import or workspace backup tools.

## Workflow features

- Summary groups have explicit form/paygrade and member selections, exception notes and frozen saved-revision snapshots. They are independent of folders. Full Navy eligibility and complete-group validation remain pending.
- Field defaults resolve command, folder ancestry, explicit group and report overrides as suggestions. Reviewed field batches retain before/after values and expected revisions, apply atomically and support guarded undo as a new batch.
- Ordered route templates copy reviewer steps into a saved-revision route. Assigned local contributors accept or return with notes. Completed, current routes support local approval locks and reasoned reopening. Revision or group changes make approval stale. Contributor profiles provide attribution, not authentication; local approval does not establish an official signature.
- Meeting comparisons freeze group revisions and allow manual ordering. Staged field proposals remain separate from drafts until explicit atomic application; stale groups and conflicting proposals are rejected.
- Five source-linked Navy draft checks retain rule-pack provenance and findings against saved revisions. Missing context and dates outside the recorded coverage require review. Promotion-capacity arithmetic has boundary tests, but is not yet a complete group eligibility/validation workflow or export certification.

## Structured PDF import

The Access adapter requires installed Adobe Acrobat automation. It recognizes the three supplied May 2025 interactive XFA templates and previews eight fields: official name, rank/rate, designator, identifier, ship/station, period start/end and official comments. Explicit identifier and date-century choices are required. Import creates a new draft with source hash, retained XML and receipt; duplicate sources are guarded. Unknown or flattened PDFs are rejected. There is no OCR, full-field mapping or merge into an existing report.

Synthetic interactive examples were filled from the blank templates. Native Acrobat data extraction and Access imports were verified. Standard PDF rendering displays the templates' fallback page, so visual form appearance remains unverified; printed layout is deferred.

## Guided workspace copies

Backup checkpoints pending editor values, closes Access, then launches a hidden Windows PowerShell worker. The worker copies the closed database exclusively and verifies SHA256, DAO reopen and table counts. Restore creates a new file and refuses overwrites. Verification JSON records success or failure. Windows PowerShell and the installed DAO engine are required; the generated worker is embedded in the app.

## Work still required

Remaining native fields need confirmed target-app mappings. Full NAVFIT import/export, stable cross-file identities, three-way returned-file reconciliation and round-trip proof remain open. Full Navy grouping/exception rules, broader versioned validation, Access Runtime/x86 and workstation verification also remain open. UX and printed layouts are deferred.

The synthetic `build/navfit-proof/NAVFIT_NEXT_FIELDS_PROBE.accdb` preserves the original TEST_DB and marks candidate qualifications, achievements, reporting-senior information and RRS fields. Target-app observations are needed before these mappings can be implemented safely.
