# AV-1 Overview and Summary

## Purpose and scope

Swift NAVFIT prepares Navy performance reports locally, for one command's reporting
senior. It covers:
- **Forms:** E1â€“E6 evaluations (NAVPERS 1616/26), Chief evaluations (1616/27) and
  officer fitness reports (1610/2).
- **Lifecycle:** draft, local review, Mark signed, and the retained history of
  every revision.
- **Groups:** summary groups with trait averages, ITA and PSG RSCA projections.
- **Counseling:** midterm counseling tracking.
- **Exchange:** NAVFIT98A import and export, structured PDF import, and
  official-form PDF output.

Out of scope: official signature, BUPERS Online submission, routing outside the
command, authentication, and network operation.

## Context

The application and its data are one Microsoft Access file, the workspace. It
runs on a standard Windows workstation with desktop Microsoft Access, from
storage the command approves for PII. There is no server, web service or network
connection. Data enters and leaves only through files the user chooses.

## Key facts

<!-- generated:facts -->
| Fact | Value |
| --- | --- |
| Application version | 0.53.0-dev |
| Schema version | 44 |
| Tables / enforced relationships | 79 / 169 |
| Report fields in the catalogue | 100 |
| VBA modules / forms | 115 / 71 |
| Validation rule packs / pack rules / expanded rules | 2 / 6 / 10 |
<!-- end generated:facts -->

## Findings and status

- **Delivered in development:**
  - preservation and history of every report revision;
  - effective-dated validation;
  - summary-group accounting;
  - counseling tracking and Mark signed;
  - automated Section 508 checks.
- **Open acceptance items:**
  - printed fit of PDF output;
  - a workstation without Acrobat;
  - screen-reader testing;
  - code signing.

  See [PROJECT_STATE.md](../../../PROJECT_STATE.md).
