# OV-6b State Transitions

The states below are values the schema enforces through table validation rules.
`test-architecture-views` checks the lists marked **checked** against the build.

## Report workflow (`Report.WorkflowState`) — checked

<!-- states:Report.WorkflowState Draft,Approved,Signed -->
```mermaid
stateDiagram-v2
  [*] --> Draft
  Draft --> Approved: local approval
  Approved --> Draft: reopen local approval
  Draft --> Signed: Mark signed
  Approved --> Signed: Mark signed
  Signed --> Draft: reopen for correction (prompted on first edit)
```

Archiving is separate from the workflow state: `IsArchived` hides a report
without changing its state, and Recover reverses it. Draft disposal (after Mark
signed) empties superseded drafts but keeps their revision rows.

## Summary group (`SummaryGroup` lifecycle) — checked

<!-- states:GroupLifecycle.WorkflowState Initialized,In-Work,Ready for Signature,Completed -->
```mermaid
stateDiagram-v2
  [*] --> Initialized
  Initialized --> In_Work: work begins
  In_Work --> Ready_for_Signature
  Ready_for_Signature --> Completed: one signed date; marks member reports Signed
  Completed --> In_Work: reopen; reopens member reports
```

## Review assignment and route step

```mermaid
stateDiagram-v2
  [*] --> Pending
  Pending --> Accepted
  Pending --> Returned
  Pending --> Cancelled
  Accepted --> Stale: saved report or group context changed
  Stale --> Pending: requester restarts or reassigns
```
