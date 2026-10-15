# r111: NAVFIT source group discovery

The supplied TEST_DB has three populated folders, but every report has blank rank/rate. Discovery previously raised on the first unresolved paygrade, preventing the entire list from appearing. Folder-default decoding could also interrupt discovery before a counseling century was selected.

Discovery now lists these source groups with `?` in the paygrade column. Select a group, choose a compatible paygrade and click **Assign paygrade to selected groups**, then Preview and Apply. Mixed selections are validated together before any assignment. Local assignments invalidate previews and may merge unresolved members into the matching folder/type/paygrade partition; Preview shows the complete selected roster. Existing valid source paygrades are preserved. Original report rank/rate values remain unchanged and can be completed after import.

Empty folders have no reports to import and receive an explicit empty-state explanation. Unsupported parent references and malformed source schemas still produce errors. Preview retains field validation, duplicate prevention, source hash checks and atomic batch commits. The supplied TEST_DB also contains an invalid reporting-senior SSN; discovering its groups does not bypass that independent save validation.

## Verification

- Original TEST_DB: all three source groups visible, original hash unchanged.
- 28 targeted checks: unresolved groups, incompatible assignment rejection, explicit compatible assignments, preview invalidation, complete atomic import, source field preservation, provenance and duplicate rejection.
- Existing 248 native group selection/import/export assertions pass.
- Both candidates compile and launch expected startup forms; blank retains 48 empty user-data tables; demo retains nine reports.
- Artifact hashes: [verification record](GROUP_DISCOVERY_R111_VERIFICATION.json).
- App version 0.25.0-dev and schema 25 retained. No CUA/manual display verification or GitHub publication.

## Candidates

- `output/ux/Swift_NAVFIT_UI_Blank_2026-10-14-r111.accdb`
- `output/ux/Swift_NAVFIT_UI_Demo_2026-10-14-r111.accdb`

Earlier builds and user workspaces preserved.

## Google Drive delivery

Uploaded on user request; connector readback confirmed exact names, sizes and destination parent. Local SHA256 hashes match the verification record; no remote checksum was returned.

- [Blank r111](https://drive.google.com/file/d/1M3EpqEN2mcG6WVj4Nnpfs2ZjnH5Jyl7u/view?usp=drivesdk): 17,719,296 bytes.
- [Synthetic demo r111](https://drive.google.com/file/d/1nKg-13w0jXA87WoBW4ZUBOIyI0g_R5nt/view?usp=drivesdk): 14,811,136 bytes.

Existing Drive files and sharing settings preserved.
