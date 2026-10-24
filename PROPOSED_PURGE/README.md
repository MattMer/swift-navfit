# Proposed purge — review holding area

Created at the user's explicit request on 21 October 2026. These files were
removed from active locations, not deleted. This is a temporary review area, not
an alternate build tree. No database, release, supplied reference or user workspace
is included. Nothing here is needed by the current manifest build. Manual use
of historical helpers remains uncertain.

Do not run these scripts in place: their relative paths assume the original
location. To restore one, verify its hash below, ensure the original destination
is absent, then move it back to that exact path. Review before final deletion.
The three assemblers deleted in the previous batch are not recreated here.

## File inventory

Paths below are both the original project-relative path and the path inside this
folder. SHA-256 values identify original bytes; every move was verified afterward.

| Original path | Bytes | SHA-256 | Reason | Discovered incoming references |
| --- | ---: | --- | --- | --- |
| scripts/build-ux-admin-feedback.ps1 | 2044 | 6f91781842789c4f7c08ddb7ebac9c3a0f86fc25e90cbf17fbd590a68554f20d | Superseded patch assembler; possible manual use. | planning/CURRENT_STATE.md, planning/MANUAL_FEEDBACK_PROGRESS_2026-10-06.md, planning/UI_AUTHOR_REVIEW_2026-10-07.md |
| scripts/build-ux-bulk-import.ps1 | 1485 | 17f839ae26298646eae7e6b4a8eb7c5cb9a4976c55b3c44c8041eb2ed375122a | Superseded patch assembler; possible manual use. | planning/BULK_IMPORT_R122_2026-10-20.md |
| scripts/build-ux-career-columns.ps1 | 968 | 65ac17d97a641738d6d0222802d24d89f0e55e1ee92b8544e4be0e0d453fadbd | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-comments-ita.ps1 | 1430 | f3248f081b30725c8816c687ae128e2597325e77990ba512933ec67ecdbfa99b | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-compat-pages.ps1 | 1456 | a0620bcbe1650e7383cfe9c01d2e817e2e242662d56488ab5febb733830d2356 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-context-help.ps1 | 1380 | 20996043ab0eeb8d1b6842fed1ba33f91b49ddb7437741a830de35abc34561a9 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-dashboard-update.ps1 | 773 | ff6b8a054488c1d9739d1744c1c70a7d702fcae90290bde4e33ee131e0b9613f | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-default-resets.ps1 | 968 | 63c78e78ab4bdfd0fa8e383bd6b068ac95045d490a0658f6f4bb124b15653458 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-directory-batch.ps1 | 969 | b42031f24bbd74b5d2188461d25a1886179c382416c77cea81eb21c55a24003a | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-directory-identity-picker.ps1 | 979 | 9083c9f732202d0283a57e6d7fe223395225fe5c3bdaaf5e424565085a42b913 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-directory-identity.ps1 | 972 | d32a8e10f051be92f98dabb0bfdb39ddcdd19ce3618718f3f69fc99c1b3847b0 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-directory-schema.ps1 | 970 | f1012c357a19c1ffa288d20d7eac4e4aa26269e351cf4e26c5173a57223e2fba | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-directory.ps1 | 963 | 18bc4b8bff59b0c850868263b50221cd3993f39eaddd92ab6336f386d661f5b6 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-export-names.ps1 | 966 | 1e6b337862b448a59ab577932ce30efd9cd9021922a84db18d7ab4e732a5ecda | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-feedback-complete.ps1 | 1718 | e23ef8091d42a1790ad9b239bc097ea249d1bff1e12fca3874f400649b60b5d3 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-feedback-tables.ps1 | 969 | 1d1c68d39ecc9e41ea67f75c5dad506e81e55107c4f5cb12a4f9a90fe2f42e3d | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-feedback-update.ps1 | 771 | 9390b84594535da53a2d6ddc815f00edef8c4bbeaf72c0d55441f67bb1d44955 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-feedback.ps1 | 962 | a8c7a2e3877a144c9bd3033c4edd713ab79f697fcbf340b5773ff32e66f78395 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-global-defaults.ps1 | 969 | 613fb691636f4cdf99fe270f333f46b63c8046732b852b55b04063a8cfb773f7 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-grid-reads.ps1 | 1055 | ab7a2abcd9dc44493a4909e5e79abc95ea5940d5d0c7414308c903e5ef190236 | Superseded patch assembler; possible manual use. | planning/PRIORITY_CLEANUP_BATCH4.md |
| scripts/build-ux-grid-update.ps1 | 773 | ab1149d67b5903469cf001707878431a155d506dca4f56fe7622beda636660bd | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-group-calculations.ps1 | 1143 | d2d9c4a3d0e6f2b8d3409630c07ded2dda6296accb17dcb365b11177d6ef77d8 | Superseded patch assembler; possible manual use. | planning/PRIORITY_CLEANUP_BATCH3.md |
| scripts/build-ux-group-correctness.ps1 | 779 | 9a33dfac0fc64aecaa1f3445f8558b3686784ae0585848e7a1dc7effd579eb48 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-group-discovery.ps1 | 809 | bc8607134ac47434f5509d3bd10c65909e1beae3a5269163088e567fc9065f2d | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-group-exit.ps1 | 1422 | 536ef0e85c9ca8bb39ddb40081f7ae6a7f79f764eca29d8a3652ac7ec3dc0c7b | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-group-organization.ps1 | 972 | 69cc5c629f1fd264874492112d446813fd513dc5430d061d45d2e7fc7f4394e4 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-group-policy.ps1 | 970 | 6e8600c53298b77193bfab93b751e6a2e2f14941f6496a9704861685afded5f1 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-header-review.ps1 | 967 | beda720fb74f6159be3eea1743b37c00c06fecf7b92a445fad3519188322faaa | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-immediate-defaults.ps1 | 816 | 81b95670126aaa40eb6af301f944188500c8cf87938383bf2ddd45eecb5479e9 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-import-progress.ps1 | 1432 | f5380bc61c185bf751653fd8f243f2172a6c8c67918d1c4349104e8b6a0a4007 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-import-rank.ps1 | 819 | cd1ea3a4fd8e48276563f5d937a45bb937e4aaa29f0c5cefb6d47cfb66d43340 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-import-scope.ps1 | 966 | 3a897d3db145881d4c211e566dcfa4a315e6e136ec6b2f7bf9b4746ee620cb1b | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-member-grid.ps1 | 965 | ca10b9be308cd6f26e581ef12a1498409ab4eb6a8aa79d2a8568e4aeda1c2b74 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-member-history.ps1 | 968 | 15ec425eac4bf965a0f6ad297a00fd1461b19f2343f060ab9253535aba05da10 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-member-psg.ps1 | 868 | 2cce1b8dd9fd45cbdab3a8dc08254f55f41f493266c0b789622f9944a53b5684 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-narrative-rates.ps1 | 969 | 6dd1d92e500cc35b9620ff7966ff17025bd36891a67c282e7aaa630111477abc | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-native-dates.ps1 | 811 | 7744160f4aa41b611b9334deea00fab40b60396c34da771efc34dc5b43527188 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-native-groups.ps1 | 967 | 90eb8c3f752f1aaea8b0a475f088f8cfe84c55ba0d27af72b3784b1a128ef01d | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-navfit-upgrade.ps1 | 923 | ec953f7d5ae6af5df5f1df38f136fb1552af56fc610a6fe880086b7bea0de69d | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-navfit.ps1 | 1254 | ea616f6bb4506d91f2a462744bda72a093b871b881b5ec8ac3dfc44adf2c6258 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-navigation-context.ps1 | 1392 | 37a7093011a4582e706c437f68ce1db9aff6111624c57d3266a5a6997e841906 | Superseded patch assembler; possible manual use. | implement_page_switch.py |
| scripts/build-ux-navywriter.ps1 | 964 | 611d64cc873a52af8aabedb00c79365a16ca258a36c560eda20d6ef08159316e | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-note-entries.ps1 | 966 | 90dd5b4d1476ac85480834aaf8b9f4c13c4f1a8d17d3de2f0abcb81f5bccbf2c | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-notes-schema.ps1 | 966 | b28d8a1f0d36382c8c7d22670a2547ecb4a5a46f24f2f6e08bcbecc937e2b2ce | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-organization-update.ps1 | 789 | d31e99bcf95a16e77727d222e7c3b42f718270ed7f46d102b5b73129e9bb53d5 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-page-switch.ps1 | 1378 | 2166f76fc6d4dcc3737d14a8cb50943b246ea05424a7b45721a05de7b1e9e1de | Superseded patch assembler; possible manual use. | implement_page_switch.py, implement_width.py |
| scripts/build-ux-page-width.ps1 | 1375 | 79d6e0d5409c52cccf47f417f669476e566460e1ad3f5ef03c2b7b2fd3e25543 | Superseded patch assembler; possible manual use. | implement_width.py |
| scripts/build-ux-pdf.ps1 | 1476 | 017e2c894d0b41d3ee9d3935255a7123f754b1cf8c71c75b5e00c2427801020b | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-pending-cells.ps1 | 1419 | 3356305de92f6dc2d5d8cab52f2eb9b8fc871721f42fa6569a69077a53364e8a | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-performance.ps1 | 1613 | 2c8d3bc988837e70a9c7c37c6d7c8a118c7cc976b87100981e1a77fc02f16e80 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-policy-expansion.ps1 | 887 | 970e0e3ad8e048512d72dce19a56412524766791957bfae7981a73cda65f29f6 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-prior-columns.ps1 | 967 | e5ff4912deaa8f8fd26cc5b6848ce89e0e61c4c5ac0d37313ef42b1fa17a4d2f | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-prior-readonly.ps1 | 1438 | bdcb190aee20e6560ac7d49ed375da2ed7938e7ab4a34b0f91d5f5daf12ee748 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-prior-tooltips.ps1 | 968 | 60524d1ceef96849511bf32d6da0a7d873958d7c18c34c5d4a355c8a7f971bf6 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-priority-actions.ps1 | 1388 | c2fe421507e913e839cc74a242c2552daf40db6cbc99f3b4b329f6967f7b4eb3 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-read-rows.ps1 | 1079 | 9010f8be2b4228ee4db825c1c67416fc9634c3051342c4bf450e114b45cc773d | Superseded patch assembler; possible manual use. | planning/PRIORITY_CLEANUP_BATCH2.md |
| scripts/build-ux-readiness.ps1 | 797 | 1fcd6c47c823b018894674bf7d126a10e6ce1d41adc9a664e15763cbdda8ad17 | Superseded patch assembler; possible manual use. | planning/READINESS_R110_2026-10-14.md |
| scripts/build-ux-report-rail.ps1 | 1296 | 43fd2fcfde08fd9def1f03f2cce9b3fb15d0682591c539a7af296c0607358ec8 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-resource-tools.ps1 | 821 | ecd99bf24b37eb4a33491050c27595ba2d53c8982d6565d680dacac276fbf0d6 | Superseded patch assembler; possible manual use. | planning/RESOURCE_TOOLS_R108_2026-10-14.md |
| scripts/build-ux-rsca-review.ps1 | 1428 | 1631a71af3c89796ff06439f54f3931ecb4f146cea644b3fc1a9d8193126d1af | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-ruad-import.ps1 | 955 | 51edaf13b47364616b97eeb0cbb171d4af36f238d2fca7c5132765bfaed4a1a1 | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-snippet-rates.ps1 | 967 | 03db9c140888a3e123e3310bed36fb76cd5b1318196a24494874f1f0b3671eee | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-theme.ps1 | 928 | 6f3e8bb4746c54ff4c268536353a43ad1ec6e14e564f1bbfdc0363251a89c1ff | Superseded patch assembler; possible manual use. | planning/SILTSTONE_THEME_R29.md |
| scripts/build-ux-validation.ps1 | 1308 | d7c6846396ae67a8d58c3c564ea2586c56a4bb60e9eb8270386258c7afab153e | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-visual.ps1 | 1368 | 0348922a648213cb0131a237185ab438c99ac7eb60463305274eefc9c74f12dc | Superseded patch assembler; possible manual use. | None found |
| scripts/build-ux-work-inclusion.ps1 | 1066 | ad998bb0ec0b0c56bc28799a5f7117c6d1a2cc7fca6fc9b158d9dfa20d25ee9a | Superseded patch assembler; possible manual use. | planning/PRIORITY_CLEANUP_BATCH5.md |
| implement_demo_groups.py | 3407 | 706d6e8152b01a41dfa3bfb78f8d94b86326121e181d02f5efc3ee5fd4327954 | One-time local helper; possible unique local work. | None found |
| implement_page_switch.py | 3174 | 2395e1ab71fe59f2d6bae2c78e05d363f9f714471b214228ee356b27a5a7ef02 | One-time local helper; possible unique local work. | None found |
| implement_width.py | 1900 | 9d5d071b98bff8dae9403f1bca2f3af5d24941e96e0fa8496c511224fbb9abf3 | One-time local helper; possible unique local work. | None found |
| measure_page_stats.py | 883 | a50bba8e58c2fac875a8a22aef150eb8e48e6989aead5a4cb0c421f1c1f1b665 | One-time local helper; possible unique local work. | None found |

Restoration does not restore an older release environment. Current source and
manifest remain authoritative; these historical wrappers may assume older schemas
or layouts. No application compatibility support was removed.

## Demo-builder consolidation

| Original path | Bytes | SHA-256 | Reason |
| --- | ---: | --- | --- |
| scripts/build-preview.ps1 | 5418 | 61cb8e5370968bd7e7ac2a441c7fcd2cc1bb45e8a92acebe3021f35558ae814c | Older preview builder duplicates seeding, lacks newer directory/history initialization and normal unload restoration; manual historical usage remains uncertain. Canonical builder is scripts/build-ux-admin-demo.ps1. Restore to original path before any reviewed use. |

## Legacy test retirement

These suites enforce superseded UI contracts and are excluded from current acceptance.
Unique persistence/import assertions are preserved here for review; complete
replacement coverage is not claimed. The tools/provenance chain stays active.

| Original path | Bytes | SHA-256 | Reason | Incoming references |
| --- | ---: | --- | --- | --- |
| scripts/test-traits.ps1 | 2233 | 70f9534b60440c70adb38d5dca26ca00842c00cdc55a2bcec1ec17b775ee6b57 | Historical combined suite; exact catalogue loop duplicated by current focused catalogue test. Mixed unique assertions retained for review. | scripts/test-foundations-suite.ps1 |
| tests/TraitsHarness.bas | 9089 | 77f482acc2ec988cc1b394979ff87bd7e73a3008ec8a7de5ef22997b51b84ad5 | Requires retired frmChecklist and fixed 78-field snapshots; unique persistence/rollback/legacy-restore assertions preserved pending port audit. | scripts/test-traits.ps1 |
| scripts/test-review-comments.ps1 | 1424 | 799c84b201a0786d3fadcd8e97d59da31b96780060e962133751d4fc3a4b7a20 | Reopens retired standalone comments UI; current embedded report-tools suite is authoritative for current conversation UI. | scripts/test-foundations-suite.ps1 |
| tests/ReviewCommentsHarness.bas | 5611 | 481b22cc5d8099b00f181d694c581e1082fbe2a65e1f3cf20b8e5fc9c4a91b6f | Requires frmReviewComments and mandatory resolve/reopen; unique legacy API assertions preserved for review. | scripts/test-review-comments.ps1 |
| scripts/test-foundations-suite.ps1 | 1377 | 0d90384c06e10a96330aac6039124d636b17f3ce2c15698412573798e3749cb3 | Historical fixed 24-suite runner calls the two retired suites; no maintained executable callers found. | None found |
| tests/GroupReadLegacy.bas | 7194 | e6301e2f56b40eb55a7e9408e1ab72e46bca2d8ee3c5a254c66e8526858eca29 | Pre-r126 group-calculation implementation used only as a legacy equivalence and timing baseline; current group tests now exercise the authoritative implementation directly. | scripts/test-ux-group-reads.ps1, scripts/measure-ux-group-reads.ps1 |
| tests/UXGridLoadLegacy.txt | 5338 | f0107befa4cb871690fdf1878aa9b0d706101e95492302e39dd0180e80c6e661 | Pre-r126 comparison-grid loader used only to compare old and current timing/output; current grid tests now verify repeatable current loads. | scripts/test-ux-grid-reads.ps1 |
| tests/UXPrepareWorkRowsLegacy.txt | 5850 | 255afc7baf6e3e375bb2641f1c32a39c07dcee65f3bea0e174e30cb20b8e2264 | Pre-r126 My Work row builder retained only for historical performance/equivalence comparison; current inclusion tests now exercise the batched builder directly. | scripts/test-ux-work-inclusion.ps1 |
| tests/UXWorkCountLegacy.txt | 818 | dabd411786c1ccf3a4a52c75e88024643dbe216cbc74c2fb41341addc3707fa0 | Pre-r126 My Work count implementation retained only for historical performance/equivalence comparison; current inclusion tests now exercise the current count directly. | scripts/test-ux-work-inclusion.ps1 |
| tests/PageTimingHarness.bas | 1054 | a613af67d5dfa7f5b2e3f6f26023219f53b72db78121e789e40a4ae8f8c6e5aa | Standalone page-switch timing harness with no current script, build, or documentation caller; superseded by the current report-open and focused page-switch measurement workflows. | None found |

## Tools/provenance dependency separation

The provenance viewer suite now owns a synthetic disposable fixture. These
files are no longer its prerequisites. Their unique snippet, reconciliation,
folder-batch and journal assertions remain preserved for review; full replacement
coverage is not claimed. No external database or production service is moved.

| Original path | Bytes | SHA-256 | Reason | Incoming references |
| --- | ---: | --- | --- | --- |
| scripts/test-tools.ps1 | 1511 | 4355a4aaeaad90cdae40911e48006851af00c14f15ed7aaf00b8b8fbe0bae995 | Integrated tools suite requires retired checklist and external TEST_DB; unique assertions held for review. | scripts/feature-check.ps1 |
| tests/ToolsHarness.bas | 10659 | 0994c31cb27dc93cb5dd7d8c0047c4205aaa1786bdda808caa1e0a8ef61853be | Integrated tools suite requires retired checklist and external TEST_DB; unique assertions held for review. | scripts/test-tools.ps1 |
| scripts/feature-check.ps1 | 1263 | 2543165e7caeb6ff9acfdfda6ca34c41f3586a154b1dcc5345805353998dd5df | Historical hard-coded 0.6.0 candidate builder; calls retired tools suite. | None found |

## Historical packaging and diagnostics

These are review candidates, not production schema/runtime removals. Current
upgrade-workspace imports src/modMigration.bas; that canonical file stays active.
Existing diagnostic databases, traces, releases and external materials stay put.

| Original path | Bytes | SHA-256 | Reason | Incoming executable references |
| --- | ---: | --- | --- | --- |
| scripts/verify-ux-review-build.ps1 | 2517 | 094e80ac84395bb37eb370e545c9d81d7c9a7355e2c6eb78a8709e099be51563 | Superseded source normalizer/verifier requires retired frmUXReview; current verify-build-artifact owns current structural checks. | None found |
| scripts/package-ux-candidate.ps1 | 9740 | 7fd640c02f8d9053d812f391e7890eb218a8473625667f97251b7f7c717d95b6 | Historical schema-23/version-0.24.0-preview packager; current pair delivery and release-check are authoritative. External/manual historical use uncertain. | None found |
| tests/modMigration.bas | 10623 | 7a485145ba9a4f28a1ee5ea24f325d88c652010efc74c002aa90b2cae716ecd4 | Stale standalone schema-25 migration copy used only by quarantined packager. Production src/modMigration.bas and upgrade-workspace remain active and unchanged. | scripts/package-ux-candidate.ps1 |
| tests/modSaveDiagnostics.bas | 317 | a7b13215fedb2c231a19a329fad442db769ae9f6ab9b19747464c44621ef9870 | Old save-trace helper with no active SaveTrace callers; manual diagnostic use uncertain. Existing diagnostic databases and logs preserved. | None found |

## Generated source-check exports

27 verification-export folders (2,567 VBA files, 30,065,890 bytes) were moved
from build/source-check-* into the same relative paths beneath this review area.
These are snapshots, not active source or databases. Their historical diagnostic
value is uncertain, so no files were deleted. Current producer generates new
unique folders; it does not read these old paths. No exact-folder references were
found in the inspected planning/docs, build/workflow/dev and root receipt metadata.

[Per-file original paths, sizes and SHA-256 inventory](SOURCE_EXPORTS_2026-10-21.json)
was recorded before relocation. Restore whole folders only to absent original
locations after checking hashes. Do not import snapshots as current source.
No databases, failure logs, release receipts or reference materials were moved.

## Browser-only prototype isolation

The native application retains prototype/fields.mjs at its existing path and
original hash. The following browser-only implementation, builder and tests are
held for review; no browser storage, generated preview output or user data was
cleared. Restore these files to their original paths before any historical use.
The builder expects the active fields.mjs in its original project location.

| Original path | Bytes | SHA-256 | Reason |
| --- | ---: | --- | --- |
| prototype/app.mjs | 74035 | 1bc6720e4c469d71ca13e46ce1d6f042bf1876c2ac94eebcef173369d7561b4e | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| prototype/model.mjs | 12450 | 0c5f9efdbd23ebd5e4931e6c81e1a4521dc2d04765b64128a5aac801df306f0c | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| prototype/traits.mjs | 26710 | b7193a4750b244ed3e880edcb98df193803826550198770ee1b8a6523e4fa9f2 | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| prototype/styles.css | 9505 | 8becfa1dc7a45ead00c907940be647795cae803f1090c3a9c25fa35406f38e31 | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| prototype/index.html | 1071 | a4107dd32e3e800019d7a41d78176a8f3bf2fc5c62d0a4924afa0ec43fbdc727 | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| scripts/build-ux-prototype.ps1 | 1241 | 1afabfb4c517b9998f9e9294c6ed4fac36f0fde7c557952ccfc092235ad7b0c8 | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| tests/ux-prototype.test.mjs | 5283 | 9d198d7ed677461b0711db01092d934ca58b41595872edeed149e4f7d3d35595 | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| tests/ux-prototype-modules.test.mjs | 2294 | 98b28d40eb305db82491bda29781eed3c4f8129f3f942814466093d7414098b6 | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| tests/ux-prototype-flow.test.mjs | 7944 | 32b28a7dbfc138a8038b0eda4894a7eefc80ef3a27696b4bec128fefe4478cab | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |
| tests/ux-rsca.test.mjs | 3448 | b6784ebdcfc1512cc71de9e311572c7f6059468cf581c894ea367bf960300964 | Separate synthetic browser implementation/test; not part of native Access runtime or manifest. |

## Trait test restored and modernized

The earlier entries for scripts/test-traits.ps1 and tests/TraitsHarness.bas are
now restoration records: those files moved back to their original paths and were
modernized in place. Their old hashes above describe the pre-modernization files,
not current active files. Git history retains those tracked originals. They are
not counted among the files still held here. Current tests preserve trait
persistence/rollback/legacy restore/recovery without retired checklist UI or
a duplicate catalogue loop.

## Batch of 24 October 2026 (Claude takeover review)

Moved at the user's request after a reference sweep: none of these is reached from
the build manifest, `develop.ps1`, `release-check.ps1`, `.vscode/tasks.json`,
`src/`, or current docs (README, AGENTS, PROJECT_STATE, CONTRIBUTING,
Development_Workflow, BUILD_PIPELINE). Only dated `planning/` notes mention some
of them. Afterwards `develop.ps1 -Action Preflight` passed (28 build-plan, 8
normalization and 837 field-definition checks).

### Tracked files

These now live at `PROPOSED_PURGE/<original path>` and are untracked; git history
still holds every one of them. Hashes were verified after each move.

| Original path | Bytes | SHA-256 | Reason |
| --- | ---: | --- | --- |
| scripts/apply-ux-narrative-rates.ps1 | 844 | f2496efc97158b30f857f0253b0ffe243944a7835a44a91f38b7f59e864c1bff | Generator stage not in build-manifest.json; never runs. |
| scripts/apply-ux-prior-tooltips.ps1 | 234 | 2c27d309533274c5cb01b519c189636579a5691300e9ca0cd9effc4b21f695dd | Generator stage not in build-manifest.json; never runs. |
| scripts/package-baseline-r98.ps1 | 5462 | 4508373cc7b5de85ac1940b3e0fbd3228ec56f5f28aedfc47f961637da872b7e | r98 one-off baseline packager. |
| scripts/verify-baseline-r98.ps1 | 2685 | b3f4b0b84adcc54ef500ed578c1b2cbdc126e2c05c77c168cf578927a2fbb3b7 | r98 one-off baseline verifier. |
| scripts/write-directory-merge-release-evidence.ps1 | 4699 | fd1968c62c7f8d2174f3ba51f0f9e2461ce151aa22f4f77d5c5700d7189333d0 | r127 one-off release-evidence writer. |
| scripts/measure-report-open.ps1 | 1840 | 02370009ec3eb7a5fc52b61113fa85f054755508dc3df649b510c56bc02c9ea1 | r114/r115 one-off timing comparison. |
| scripts/create-scale-fixture.ps1 | 1849 | 55b4f42aecb783e41e73ab4dab67deb0c623023e60675dea33dfc80636e0d08f | r102 scale fixture; superseded by measure-ux-read-rows.ps1/test-batched-reads.ps1 seeding. |
| scripts/profile-scale-load.ps1 | 1832 | 1c1ff86cd0cc9f6fa77223a6e7d7a71aa85312d628bb1fa65ca44494f227c890 | r102-era profiling one-off; superseded by batched-read benchmarks. |
| scripts/measure-scale-groups.ps1 | 4291 | c7beebe2c64b91c2073a70be403ee5d1e1049f9e83bdd29cef673adb180d38e9 | Earlier group-scale benchmark; superseded by test-batched-reads.ps1 and test-ux-group-lifecycle.ps1 -GroupListBenchmark. |
| scripts/measure-ux-group-reads.ps1 | 811 | e454dc1b74c23fdc1e1821455f8f0345dde28da76e24eef93e0f66a909eeaf12 | Earlier group-read benchmark; superseded as above. |
| scripts/seed-directory-history-demo.ps1 | 707 | 7b67d962afa6da88cb9fab3dea7c5ac470b47deb94a8c99d9c22b30b0048aab8 | Ad hoc demo seeder; build-ux-admin-demo.ps1 is the single supported Demo builder. |
| scripts/seed-ruad-profile-demo.ps1 | 866 | f872cd67cbc3d2b48200d7c3ab1b01428c10266eb253a2d0f55fa0c69eaf3c5b | Ad hoc demo seeder; build-ux-admin-demo.ps1 is the single supported Demo builder. |
| tests/UXRUADSeed.bas | 3022 | a63aa7ecf6bf23ac183bb8108a1e984774cf4143c43e6536a32da21a543d2f68 | Only loaded by seed-ruad-profile-demo.ps1 (moved). |
| scripts/pdf-structured-preview.py | 4409 | 5a2a53860b5e36b677e4e13d6528d7708c532834df3b74de583720eedf7cc2d3 | Pre-native Python XFA preview; superseded by native modPDFLocal/modPDFImport. |
| scripts/test-access-mcp.py | 3727 | 30b2eb2efb1f2615e3e33eb8ae1f150e166d8d3493f900103ddae93ffd6c7013 | Codex-era Access Explorer MCP setup smoke test. |
| scripts/verify-installed-source.py | 839 | 58cc73e629c46e9b350e8d6cfe67b3cf3a6b6394b811b65c978850f144529331 | Superseded by verify-built-application.ps1 installed-source verification in every build. |
| scripts/inspect-navfit.ps1 | 3852 | 599fbefbc9ced51ead25dac3a97b17c0252646c2a372ad8ed6986965b892f33d | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/inspect-navfit-bindings.ps1 | 3297 | 398fc042d914a697414f0fa697eab5509d115960730b3d3d3fb9a3f2117ceb6f | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/inspect-navfit-narrative-layout.ps1 | 3389 | 8163f71a8f2427e41b03928369563f0fc07a07da7dae8108d4929c69fece5cd7 | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/inspect-navfit-print-export.ps1 | 3150 | 06410e08bb8311e3c5869a38c7f29b2c650cecd74e2824d0a5351c65b8e53541 | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/inspect-navfit-print-rea.ps1 | 1887 | ac163b25fcfe55d389cdc1a3beb66be1f2165bd71d5eda9a5c3973743b7df27f | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/make-navfit-fields-probe.ps1 | 4577 | 957396d491aa531c7fb9b868d792c9c952a02dd74642fb34cfa58a546af770d1 | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/make-navfit-mapping-probe.ps1 | 3706 | 015ac5f0e8bdca9073e886f55cd0ac86ca270eda5f821313cd5f9ed88d8e8050 | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/prove-navfit-drawing.ps1 | 3515 | 3b03038c70d87b3b5633cfc6b8a77e90f76db1d9f2b90da1ce0aaa1549e1cf00 | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/prove-navfit-runtime.ps1 | 5107 | 256630d54d4e9db06a932da3839913b5e53018e5d11257fe1e78c4fa0dd2386f | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/stage-navfit-proof.ps1 | 6107 | b7669ba98c40371d5983aa8ddf2cecc21a3a696ba8699305de5421569b453c7c | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/summarize-navfit-print-rea.cjs | 3135 | d16a0edf248fdab2d6eed20c0ea155d9c2a6df6382dab6a938dadccb37d81f59 | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |
| scripts/summarize-navfit-validation-rea.cjs | 1983 | 11a6301120f5dfca086e2b4b2374033680ec3f637e12429a1efa695c3ea021ea | Completed NAVFIT98A reverse-engineering/proof tool; evidence retained in planning/. Restore if the native format changes. |

### Untracked scratch (68.7 GB)

These moved to `PROPOSED_PURGE/build-2026-10-24/<original path>`. The per-item
inventory (paths, sizes, reasons, and SHA-256 for single files under 200 MB) is
in `PROPOSED_PURGE/build-2026-10-24/INVENTORY.json`. Large directories were not
hashed. They are disposable outputs, not inputs.

| Original location | Items | Files | MB | Reason |
| --- | ---: | ---: | ---: | --- |
| build/tests/* | 3781 | 12144 | 66636 | Per-run disposable database copies and receipts written by test runners. The two fixture inputs `pdf-counseling-status-v2` and `pdf-local-expanded` were kept. |
| build/* loose files | 863 | 863 | 3177 | Logs, probes, one-off helpers, PR/release drafts and 0.1.0/r122 scratch databases from earlier sessions. |
| build/dev/* | 14 | 14 | 186 | Dev builds without a build receipt (intermediate or diagnostic) and the superseded 0.30/0.37 unsigned ACCDE candidates. Receipted builds, which serve as per-schema legacy fixtures, and the 0.38 ACCDE were kept. |
| build/source-check-*, source-update-* | 65 | 7058 | 91 | Source-comparison scratch directories. |
| build/ui-review-2026-10-07, tools, research, ux-reference, deployment-0.23.0, rea-navfit-*, navfit-proof, navfit-staging* | 14 | 761 | 216 | Research and review scratch; no active script reads them. |
| build/test-dependencies | 1 | 687 | 37 | Duplicate of `build/dev-pdf-libs` with an older Pillow. |
| .venv | 1 | 894 | 12 | Empty Python 3.15 virtual environment (pip only). |
| scripts/\_\_pycache\_\_ | 1 | 1 | 0 | Python bytecode cache. |

Still in place: `build/dev` receipted builds, `build/workflow` receipts,
`build/dev-pdf-libs`, the two PDF fixture folders, all of `output/`, and the
supplied reference folders. Deleting `PROPOSED_PURGE/build-2026-10-24/` frees
about 69 GB once the review is finished.
