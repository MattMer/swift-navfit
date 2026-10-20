# Swift NAVFIT release requirements

New releases must import data from every previously supported Swift NAVFIT schema. Preserve all source rows, field values, IDs/relationships, revision/review history, provenance, recovery, command members, notes, defaults and reporting-senior history. Use the production migration engine and Settings & Help import, read source databases through read-only DAO without executing source code, and verify copied values before committing. Existing populated workspaces must never be silently replaced or merged by ID.

Keep schema evolution additive and versioned. Before publishing a release, run native data-transfer checks from retained earlier release fixtures into a new blank build, plus rollback/unsupported-source guards. Document supported schemas and any unverified old fixture; do not claim untested compatibility. Verify blank distributions remain blank and synthetic demos contain only synthetic examples. Preserve user-edited workspaces and earlier releases.

Report editor fields use fixed Front Page/Back Page geometry based on NAVPERS block order and official captions. Keep all seven performance traits together on Front Page as requested by the user; this editing grouping does not change printed page associations. Keep report content to the right of the application navigation rail. Window changes may adjust the viewing area, not reflow or resize report fields. Preserve the unbound buffer, saves/revisions, pending values and recovery across page changes. Do not imply verified printed fit.

## Delivery and cleanup

Place the current Blank and Demo artifacts directly in `D:\Development\Swift NAVFIT\output`, using distinct dated/versioned names and retaining earlier files. Upload the verified Blank and synthetic Demo to the user's Swift NAVFIT Google Drive folder: https://drive.google.com/drive/folders/14vylyonimhXhhJzaQhh0c51_UenHhzGg . Verify names, sizes and folder membership after upload. Complete applicable release migration and rollback gates before publishing.

When finished working, close Microsoft Access, including owned automation/preview sessions. Preserve pending user edits through the normal save/recovery workflow; do not force-kill an instance holding unsaved work.
