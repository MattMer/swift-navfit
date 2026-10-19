# r114 — import native dates containing times

## Fix

NAVFIT stores six report fields and five folder-default fields as Access Date/Time. Values can carry a non-midnight time even when NAVFIT displays only their date. Swift rejected these in NAVFITDecode before any report/group could import.

Import now uses the typed value's calendar date without a timezone shift or date rollover. Nulls stay null, non-date input still fails, and the original numeric Date/Time value remains in ImportRawValue for provenance. Preview explains the normalization. Export remains strict about unexpected times in Swift's own date-only fields.

## Temporary workaround for r113 and earlier

Close NAVFIT98A. Open a copy of the native database in Microsoft Access. In Create > Query Design, close Add Tables if shown, choose SQL View, paste and run the first query. Then run the second query separately. These updates remove stored clock values and retain blank dates. Close Access and select the cleaned copy for import. No cleanup is required when using r114.

```sql
UPDATE Reports SET
  FromDate = Int([FromDate]),
  ToDate = Int([ToDate]),
  DateReported = Int([DateReported]),
  PhysicalReadinessDt = Int([PhysicalReadinessDt]),
  RaterDate = Int([RaterDate]),
  SeniorRaterDate = Int([SeniorRaterDate]);
```

```sql
UPDATE Folders SET
  FromDate = Int([FromDate]),
  ToDate = Int([ToDate]),
  DateReported = Int([DateReported]),
  RaterDate = Int([RaterDate]),
  SeniorRaterDate = Int([SeniorRaterDate]);
```

The Folders table does not contain PhysicalReadinessDt. These statements were executed against a separate synthetic native copy and verified to remove times while retaining null dates.

## Verification and delivery

- 63 targeted timestamp import/provenance assertions covering all three report types and dated folder defaults; source SHA256 unchanged.
- 507 native adapter, 410 exchange, 99 stale/hash/identity/lock/undo guard and 248 atomic group import assertions passed (1,327 total).
- Both final candidates compile/start, blank has 48 checked empty user-data tables, demo has nine synthetic reports. Schema 25/app 0.25.0-dev retained. No manual display or actual NAVFIT printing claim.
- [Verification manifest](NATIVE_DATES_R114_VERIFICATION.json).
- [Blank r114](https://drive.google.com/file/d/1FVhtbDlvbUGmcTqHZbNPLk6cBBYwCAdO/view?usp=drivesdk).
- [Synthetic demo r114](https://drive.google.com/file/d/198JFSwUYwYfRviM7MiDkO3cf4OrB-NtY/view?usp=drivesdk).

Drive metadata readback confirmed exact names, byte sizes and destination folder. No remote checksum was returned. Existing workspaces and earlier releases remain untouched. GitHub remains at the r113 source checkpoint pending a requested push.
