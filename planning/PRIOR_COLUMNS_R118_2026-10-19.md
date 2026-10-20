# Prior evaluation reference cells and delivery - r118

Prior ITA and Prior promotion in the summary-group member grid are grey, locked, disabled, and excluded from tab navigation. Current traits, promotion and career fields retain their normal editing behavior. The shared visual palette preserves the reference-cell shading after activation. The column headings explain that these are historical reference values.

Returning from an edited member report now performs an authoritative check of that selected row before painting its locks. Ordinary row-focus events retain their cached fast path. This restores visible stale-row locks without whole-group queries on ordinary navigation or weakening the write guard.

Current artifacts are directly in `output/`: `Swift_NAVFIT_Blank_2026-10-19-r118.accdb` and `Swift_NAVFIT_Demo_2026-10-19-r118.accdb`. Earlier files are preserved; r117 was an unpublished intermediate candidate. The Blank has no user-data rows; the synthetic Demo has 13 reports, 5 groups, and 12 prior evaluations. The prior RSCA completion and Review fixes are included.

AGENTS.md now records the requested output location, Google Drive destination and requirement to close Access after work. The destination is https://drive.google.com/drive/folders/14vylyonimhXhhJzaQhh0c51_UenHhzGg . File hashes, native checks and verified upload identities are recorded in the accompanying verification manifest.

Verification passed: 17 prior-grid checks, 12 member-grid checks, 107 palette/navigation checks, 39 demo checks, 40 lifecycle checks plus forced-shutdown save/recovery, and 48 smoke checks. Transfers passed for all 25 retained schemas and for populated schema-20 and the current 13-report demo; exact original rows/fields, source hashes, reopen, rollback and rejection guards were verified. Both Drive uploads were read back with matching names, sizes and the requested parent folder; the connector did not return a remote checksum. Access is closed.
