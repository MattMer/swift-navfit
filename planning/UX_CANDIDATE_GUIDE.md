# Swift NAVFIT deployment prerelease

This package is **0.23.0-preview, schema 21**. It is a deployment prerelease,
not a replacement for the distributed 0.21.0 checkpoint. Manual physical-display,
keyboard and screen-reader acceptance remains outstanding. This prerelease includes Siltstone colors, reduced tab-switch layout work, live
validation with policy references, expanded PDF import and NAVFIT exchange.
The packaging script creates local assets; publication is a separate step.

## Choose the right file

- **Swift_NAVFIT_0.23.0-preview-DEMO.accdb**: nine synthetic reports and three
  groups for experimentation. Do not transfer operational data into this file.
- **Swift_NAVFIT_0.23.0-preview-blank.accdb**: a blank, separate workspace.
  Complete command/contributor setup or migrate into a new destination.

Use desktop Microsoft Access. When active content is blocked, follow the
instruction screen and reopen after enabling content. The build does not change
Office trust settings. Runtime and x86 remain unverified.

The in-app guide is at **Settings & Help > Workspace guide and quick start**.
The same workflows are in `guides/UX_WORKSPACE_GUIDE.md`. `MANIFEST.json` records
hashes, blank-table counts, checks and remaining acceptance work. Verification
logs and installed source exports are in `verification/`. The packager also
creates a ZIP and a separate `.zip.sha256` file for sharing; earlier candidates
are preserved.

## Transfer an existing workspace

Close the existing file in Access. Keep the original workspace. From the package
folder run:

```powershell
powershell.exe -NoProfile -File .\scripts\upgrade-workspace.ps1 `
  -SourcePath 'C:\path\existing.accdb' `
  -TemplatePath '.\Swift_NAVFIT_0.23.0-preview-blank.accdb' `
  -DestinationPath 'C:\path\upgraded-review.accdb'
```

The source is opened read-only through DAO, not launched as an Access app.
The transfer requires a blank template, rejects unknown source tables and
existing destinations, copies within a transaction, and compares every original
row and field. The source hash is checked; the destination is reopened before
success is reported. A `.migration.json` companion records the result. New
schema-21 senior/baseline tables remain empty for schema-20 sources.

After migration, open the new destination and inspect report names, periods,
traits, narratives, history, groups and review work before adopting it. Migration
fixtures demonstrate their tested scope; they do not certify every operational
workspace. Continue using the original if review exposes a problem. Opening the
original file is the rollback; migration does not overwrite it.

## Verification and sharing

Compare a file's SHA256 with its entry in MANIFEST.json before opening. Access
may change the file hash after use. Share the blank candidate or synthetic demo;
actual operational workspaces can contain report information and should be
handled through your normal collaboration process. Local copies do not sync.

Printing, full NAVFIT export/round-trip interoperability and complete Navy policy
validation are not established by this package. Local approval is not an
official signature. No network service, authentication or synchronization is added.
