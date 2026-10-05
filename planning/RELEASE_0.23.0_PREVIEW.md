# Swift NAVFIT 0.23.0-preview

Deployment prerelease, schema 21. Includes a blank workspace and a populated
synthetic demo with nine reports, three summary groups, review work, reporting
seniors/RSCA and recovery examples. Original workspaces and older builds remain.

## Changes since 0.21.0

- Native workspace redesign: persistent navigation, official-block sections,
  denser forms, trait comparison, summary-group counts and reporting-senior tools.
- Siltstone palette across 51 forms and 1,717 controls.
- Faster workspace navigation: layout skips unchanged dimensions; context refresh
  sizes only the active workspace; redundant list requeries removed.
- Unified My Work with ordered route steps, stale-review actions and unresolved
  comment threads; contextual review and recovery safeguards.
- Live report validation and group-wide findings with specific policy references
  where mapped; these checks do not represent complete Navy validation.
- Expanded structured PDF mappings, existing-draft merges and extraction checks
  for target computers; flattened PDFs require manual entry.
- Supported-value NAVFIT import, report/group export, retained native provenance
  and single-report three-way returned-file conflict handling.

## Use

Extract the ZIP. Open the DEMO database to explore synthetic data. Use the blank
database for a new workspace, or transfer existing schema-20/21 data to a new
blank destination using the included upgrade script. START_HERE.md explains the
transfer. Never migrate operational records into the seeded demo.

## Verification and limits

The exact deployment package has verification logs and a SHA256 file manifest.
Native source audit compares every installed standard module, checks editor and
comparison unload guards, and compiles. Blank table counts and schema-20/21
source-preserving transfer are checked during packaging.

The r30 performance milestone passed 1,801 layout/focus/identity assertions,
including unchanged-size skips, real resize handling, contextual invalidation
and hidden-tab isolation; 36 demo checks passed. Actual tab-switch timings have
not been measured. Prior r27/r28 evidence covers the documented NAVFIT mapping,
round-trip and group-export scope; see the linked milestone notes.

Manual physical-display/DPI, keyboard and screen-reader acceptance remains
outstanding. Access Runtime and x86 are unverified. Returned summary-group files,
complete Navy policy validation, printed fit and direct printing from Swift are
pending. Exported files can be opened and printed in NAVFIT98A; no direct print
bridge is included. Contributor profiles provide attribution, not authentication;
separate copies do not synchronize.

Deployment packaging passed all 13 packaged suites, schema-20/21 transfer and
demo verification. All 57 installed modules match source; all blank user tables
are empty. See DEPLOYMENT_0.23.0_VERIFICATION.json and the package MANIFEST.json.

