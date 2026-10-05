# Trait editor

> Current native navigation: use the [workspace guide](UX_WORKSPACE_GUIDE.md) and
> [current state](CURRENT_STATE.md). This guide retains service details and
> earlier verification history; legacy screen instructions may differ.

Swift NAVFIT 0.10.0 adds manual marks for blocks 33–39 in EVAL, CHIEFEVAL and FITREP. Captions and general/selected-mark descriptions come from the supplied NAVFIT98A 2.2.0.332 catalogue, embedded in the app.

Open a draft and scroll to Performance traits. Choose NOB or 1.0–5.0, or clear a mark. Blank remains distinct from NOB. Incomplete drafts can be saved; unsupported marks are rejected without partial saves. Changing a new draft's type with pending marks asks whether to clear them because meanings differ.

Marks participate in snapshots, comparison, restore, duplication, checkpoint recovery and stale-review checks. Legacy unrecorded traits restore to null; comparison displays Not recorded. Raw pending input can be recovered even when invalid for saving. The checklist flags missing marks and NOB for review.

Trait averages, whole-report observation status, recommendations, summary-group rules and Navy validation remain pending. External trait codes have not been mapped or round-trip verified; imports remain narrative-only. Printed layout is deferred.

Verification: 75 trait assertions, 169 exact catalogue checks and the existing 350 assertions passed. Narrative import verified all 2472 retained values. Schema 9-to-10, 10-to-10 and 5-to-10 migrations preserved every original row/field and source hash, including checkpoints, and reopened. Runtime/x86 and comprehensive visual testing remain pending.
