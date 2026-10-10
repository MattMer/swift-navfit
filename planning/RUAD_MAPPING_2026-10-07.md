# RUAD sample mapping — 7 October 2026

Received `87114_SMART_RUAD - 1OCT2026 - SPLIT.csv` from Downloads. Read-only CSV inspection found 35 columns, 66 aligned rows, 60 named members and six vacant billets. No malformed rows or repeated nonempty names in this synthetic sample. Values need whitespace trimming (168 padded cells). All identifiers must remain text, including UIC/BIN and codes.

| Source | Proposed use |
|---|---|
| NAME | Member display name and suggested official report name; preserve punctuation and spelling, do not truncate |
| DESG for enlisted | Actual member rate (e.g. IT2), distinct from billet requirement |
| RATE/RANK for officers | Officer rank (e.g. LT); retain DESG separately for designator review |
| E/O | Officer/enlisted classification for mapping and report-type suggestions |
| RUIC / RESERVE UNIT | Selected roster/unit context; not an automatic assignment to report Block 6 |
| SELRES TYPE | Roster assignment context; Vacant rows do not create people |
| BIN | Billet provenance only; never a permanent member identity |
| REQ DESG/REQ RTING | Billet requirement; must not overwrite a member's actual rate |
| DESG for officers | Raw officer designation has five-character values with trailing letters; require an explicit preview rule before filling four-character report designators |
| PRD | Optional personnel context after date parsing; never the report ending date |
| ADDRESS / PHONE / NROWS Email | Retained in the member profile Contact page; sample values are redacted placeholders |
| Other mobilization/status codes | Retained as roster metadata in the member profile, separate from report fields |

No SSN is supplied. New members can be provisional, but report export still needs SSN entered elsewhere. Existing directory records must be matched through a user-reviewed choice; names and BIN are not proof of identity. RUAD refresh must not overwrite SSN, prior evaluation history, report links, individual report fields or manual profile changes. Missing members in a later file require review rather than automatic deletion/archive.

Next implementation: Choose → Preview → Apply → Results, skip six vacant rows with explanation, show member rate/rank and unit context, offer explicit existing-member matches, detect ambiguous/repeated selections, and apply atomically with recorded provenance. Add an additive source/import-batch model only if needed to support repeat refresh and undo; current CommandMember has no roster-source key or source-row metadata. The sample has now been received; earlier notes saying sample mapping is blocked are historical.

The r90 member profile now exposes all 35 source columns across Unit, Member, Assignment, Mobilization and Contact pages. NAME uses the existing member display-name field. The separate report rank/rate remains distinct from raw RUAD RATE/RANK and DESG. Schema 25 adds one nullable CommandMember.RUADData memo field; profile updates save the full payload atomically with version checks. Existing profiles remain valid with no payload. Two clearly named synthetic profiles in the demo exercise the complete fields. Full CSV import/refresh is still pending; these profiles do not establish an import workflow. Analysis counts are recorded in RUAD_SAMPLE_ANALYSIS_2026-10-07.json.

## r97 implementation

The reviewed import/refresh workflow is implemented and verified with 71 assertions; see RUAD_IMPORT_R97_2026-10-09.md. Earlier pending-import notes above describe historical states. No schema beyond 25 is introduced.
