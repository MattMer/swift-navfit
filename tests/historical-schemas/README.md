# Historical schema definitions

One file per Swift NAVFIT schema older than the current build. Since 10 November
2026 the importer accepts schema 35 (release 0.49.0) onward, and the import test
builds only the schemas listed in [release-schemas.json](release-schemas.json)
plus the current one. The older files stay as the record of what shipped;
release 0.55.0 was the last to import them.
[test-swift-data-import.ps1](../../scripts/test-swift-data-import.ps1) builds each
into a data-only database, seeds synthetic rows, imports it through the
production Settings import and compares every row and field. Old release files
are not needed.

Each file starts with `--` header lines, followed by one DDL statement per line:

| Header | Meaning |
| --- | --- |
| `schema`, `variant` | Schema number. A variant letter marks a schema number that shipped with more than one structure. |
| `kind` | `git-ddl`: exact DDL from the git commit at that version. `captured`: DDL generated from a read-only capture of a real build. `reconstructed`: schemas 1–7, which predate git, rebuilt from schema 8 minus the documented differences listed in `note`. |
| `source`, `release`, `appVersion` | Where the definition came from. |
| `zeroLengthOff` | Optional text fields that do not allow zero-length strings. Every other optional text field does. |
| `validationRule` | Table validation rule, as `<table> <rule>`. |

Before raising `SCHEMA_VERSION`, capture the outgoing schema from a build of it:

```powershell
powershell.exe -NoProfile -File scripts/capture-schema-definition.ps1 -BuildPath <build>
```

Do not edit captured or git-derived files by hand. They record what shipped.
