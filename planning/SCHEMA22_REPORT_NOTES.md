# Schema 22 - attributed report notes

The requested contributor/note entry list cannot be represented faithfully by
one DraftNotes memo or by review-comment threads without changing their meaning.
Schema 22 adds only ReportNote and ReportNoteDraft. No original table/field is
removed or rewritten. Directory and prior-eval storage are not part of this
increment and will need a separate additive design after RUAD mapping is known.

Posted notes are append-only, record the contributor ID and name at posting,
and remain separate from NAVPERS comments, approval and report revisions.
Unposted notes have unique identities, author/report binding and optimistic
checkpoint versions. Explicit discard cannot delete another window's newer
checkpoint. Only the owning contributor can recover or post a checkpoint.

Legacy DraftNotes is retained verbatim in reports and revisions. An idempotent
migration creates a separate entry labeled "Legacy notes (author not tracked)";
it does not infer an author from the most recent report editor. A unique nullable
legacy-report identity prevents repeat conversion.

Migration supports existing schemas through 22, preserves source hashes and
verifies every original row/field before commit. Schema-22 notes/checkpoints
must also survive a schema-22 transfer. Schema downgrade is rejected. Published
schema-21 artifacts are preserved. These development builds retain the existing
app version string; they are not a new published deployment.

Checks so far: 19 notes/checkpoint service assertions, 18 notes/recovery UI
assertions, and schema-21 demo -> 22 transfer with all original rows/fields and
nine legacy note entries verified. Final UI and schema-22 transfer verification
remain in progress.
