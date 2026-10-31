# Swift NAVFIT architecture description

A fit-for-purpose architecture description of Swift NAVFIT that follows the
DoD Architecture Framework (DoDAF) 2.02. Its structure follows ISO/IEC/IEEE 42010
(*Systems and software engineering — Architecture description*).

- **Conformance:** this description is **DoDAF-aligned, not DoDAF-conformant**.
  DoDAF conformance needs data defined in the DoDAF Meta Model (DM2) and
  exchanged in its Physical Exchange Specification (PES), which takes a DoD
  architecture tool. The views here are documents and tables built from the
  application's own sources. Section 2.4 explains this choice.
- **Generated views:** views marked *generated* come from tracked source and a
  fresh build through `scripts/generate-architecture-views.ps1`. The
  `test-architecture-views` runner fails when a generated view no longer
  matches the build.
- **Current state:** [PROJECT_STATE.md](../../PROJECT_STATE.md) holds current
  behavior and decisions. This document doesn't repeat it; it organizes it for
  architecture review.

## 1. Introduction

### 1.1 Identifying information

| Item | Value |
| --- | --- |
| Architecture | Swift NAVFIT local reporting-senior workspace |
| System of interest | Swift NAVFIT, a Microsoft Access desktop application, with its workspace file |
| Version described | Development source; the application and schema versions are in [AV-1](views/AV-1.md) |
| Architecture type | Solution architecture for one standalone application (DoDAF "solution" level) |
| Classification | Unclassified. The application processes PII (CUI) on command-approved storage; this document contains no personnel data |

### 1.2 Supplementary information

- **Owner:** the Swift NAVFIT maintainers. The repository README names the
  current contact.
- **Change control:** git history. Each change to a hand-written view is
  reviewed like any other source change.
- **Reference documents:**
  - DoDAF 2.02;
  - ISO/IEC/IEEE 42010:2011 and the Hilliard architecture-description template;
  - DoD CIO memo "Software Development and Open Source Software" (24 Jan 2022);
  - DoD Enterprise DevSecOps Activities & Tools Guidebook tables;
  - BUPERSINST 1610.10H;
  - Section 508 (36 CFR 1194) with WCAG 2.0 AA.

  Local copies are kept in `FORMS_REFERENCES_TEMPLATES/` (not in the repository).

### 1.3 Other information

#### Overview (AV-1 summary)

Swift NAVFIT helps a command's administrative staff prepare Navy enlisted
evaluations, Chief evaluations and officer fitness reports (NAVPERS 1616/26,
1616/27 and 1610/2) locally. It also:

- organizes reports into summary groups and computes ITA, group average and
  PSG RSCA projections;
- runs local review routes;
- tracks midterm counseling and records Mark signed;
- exchanges data with NAVFIT98A and the official PDF forms.

The workspace file is the command's **reporting-senior binder**, kept on storage
the command approves for PII. Swift does not sign, submit or route reports
officially. Those steps stay with BUPERS Online and the official process.

The full AV-1 is [views/AV-1.md](views/AV-1.md).

#### Architecture evaluations

| Evaluation | Result | Where |
| --- | --- | --- |
| Section 508 / WCAG 2.0 AA (automated) | 12 contrast pairs at 4.5:1 or better; every input and button named; Windows contrast themes followed | `scripts/test-workspace-presentation.ps1`; conformance report draft |
| Section 508 (screen reader, attended) | Not yet performed | [Screen-reader test](../Accessibility_Screen_Reader_Test.md) |
| Historical import (every prior schema) | Every schema 1–34 imports through the production path | `scripts/test-workspace-migration.ps1` |
| Extended regression tier | All runners pass on the current build | `scripts/develop.ps1 -Action ExtendedTest` |
| Printed fit of PDF output | Not yet performed | Backlog H01 |
| DevSecOps practice | 55 activities done, 29 partial, 0 gaps, 14 command-owned, 20 not applicable | [DevSecOps alignment matrix](DevSecOps_Alignment.md) |
| Malformed file input | 55 of 55 malformed and control inputs handled cleanly | `test-external-import` |
| Threat model | STRIDE over 5 assets, 5 trust boundaries and 6 interfaces; 6 recommendations | [Threat model](Threat_Model.md) |
| Software bill of materials | Generated from the build; no third-party runtime libraries | [SBOM](SBOM.md), [CycloneDX](sbom.cdx.json) |

#### Rationale for key decisions

| Decision | Why |
| --- | --- |
| A single-file Microsoft Access application | Runs on standard Navy workstations with Office and needs no server or network accreditation. Admin staff can move or back it up. |
| Additive, versioned schema; every release imports every earlier schema | The workspace is a long-lived binder, and no upgrade may lose report history (a CONTRIBUTING.md product invariant). |
| Unbound report editor with recovery checkpoints | Pending edits survive failures and restarts without half-saved revisions. |
| File interfaces only (NAVFIT98A, PDF, CSV) | Matches how reports actually reach official systems today. Adds no network attack surface. |
| Local contributor attribution, not authentication | Access controls come from the command's folder permissions (decision C01). Swift records who did what; it does not claim to authenticate. |

## 2. Stakeholders and concerns

### 2.1 Stakeholders

| Stakeholder | Role |
| --- | --- |
| Reporting senior | Owns the binder and the reports. Signs official copies outside Swift. |
| Command administrative staff | Prepare drafts and groups; track counseling; run review routes; export. |
| Contributors and local reviewers | Edit, comment on and review reports, attributed by contributor profile. |
| Rated members | Data subjects. They never use Swift directly. |
| ISSM / command cyber staff | Approve installation, PII storage location, macro trust and code signing. Assess cyber risk. |
| Maintainers | Develop, test and release Swift NAVFIT. |
| Testers | Run the Demo builds and report problems. |

### 2.2 Concerns

| ID | Concern |
| --- | --- |
| C1 | Reports are correct for the official forms and the BUPERSINST 1610.10H rules in effect for their period. |
| C2 | No report data or history is ever lost, across saves, failures, upgrades and imports. |
| C3 | PII stays on approved storage and appears only where needed. |
| C4 | Interoperability with NAVFIT98A and the official PDF forms. |
| C5 | Usability and accessibility (Section 508) for administrative staff. |
| C6 | Runs on a standard Navy workstation without servers, network services or extra software. |
| C7 | Supportability: diagnosable errors, verified backups, repeatable builds and tests. |
| C8 | Secure supply chain and trust: few dependencies, signed releases (planned), a known component list. |

### 2.3 Concern–stakeholder traceability

| Concern | Reporting senior | Admin staff | Contributors | Rated members | ISSM | Maintainers | Testers |
| --- | --- | --- | --- | --- | --- | --- | --- |
| C1 Correct forms and rules | ● | ● | ● | ● |  | ● | ● |
| C2 No data loss | ● | ● |  | ● |  | ● |  |
| C3 PII protection | ● | ● |  | ● | ● | ● |  |
| C4 Interoperability | ● | ● |  |  |  | ● | ● |
| C5 Usability, accessibility |  | ● | ● |  |  | ● | ● |
| C6 Standard workstation |  | ● |  |  | ● | ● |  |
| C7 Supportability |  | ● |  |  | ● | ● | ● |
| C8 Supply chain and trust |  |  |  |  | ● | ● |  |

### 2.4 Viewpoints selected, and those not used

DoDAF 2.02 says to choose models that answer the stakeholders' questions. These
were chosen, with the concerns each one frames:

| Viewpoint | Models used | Concerns |
| --- | --- | --- |
| All | AV-1 Overview and Summary; AV-2 Integrated Dictionary (*generated*) | all |
| Operational | OV-1 High-Level Operational Concept; OV-4 Organizational Relationships; OV-5b Operational Activity Model; OV-6a Operational Rules (*generated*); OV-6b State Transitions | C1, C2, C5 |
| Data and Information | DIV-2 Logical Data Model and DIV-3 Physical Data Model (*generated*) | C2, C3, C7 |
| Systems | SV-1 Systems Interface Description; SV-4 Systems Functionality; SV-7 Systems Measures; SV-8 Systems Evolution (*generated*) | C4, C6, C7 |
| Standards | StdV-1 Standards Profile | C1, C5, C8 |

Not used, and why:

- **Services Viewpoint (SvcV-1 to SvcV-10):** Swift provides and consumes no
  network services.
- **Capability Viewpoint (CV-2 to CV-7):** most of these support acquisition
  portfolios. Swift is a command-level tool, not a program of record. CV-1
  (vision) is covered by AV-1.
- **Project Viewpoint (PV-1, PV-3):** there is no acquisition project structure.
  The schedule lives in the [priority backlog](../../planning/PRIORITY_BACKLOG.md).
- **OV-2 / OV-3, SV-3, SV-5 and SV-6 (resource-flow models):** there is one
  system with file exchanges only, all shown in SV-1.
- **DM2 / PES exchange:** needs a DoD architecture tool (see the conformance
  note at the top).

## 3. Viewpoints and views

Each view is a separate file under [views/](views/):

| View | File | Kind |
| --- | --- | --- |
| AV-1 Overview and Summary | [AV-1.md](views/AV-1.md) | written |
| AV-2 Integrated Dictionary | [AV-2.md](views/AV-2.md) | generated |
| OV-1 High-Level Operational Concept | [OV-1.md](views/OV-1.md) | written |
| OV-4 Organizational Relationships | [OV-4.md](views/OV-4.md) | written |
| OV-5b Operational Activity Model | [OV-5b.md](views/OV-5b.md) | written |
| OV-6a Operational Rules | [OV-6a.md](views/OV-6a.md) | generated |
| OV-6b State Transitions | [OV-6b.md](views/OV-6b.md) | written, with states checked against the schema |
| DIV-2 Logical Data Model | [DIV-2.md](views/DIV-2.md) | generated |
| DIV-3 Physical Data Model | [DIV-3.md](views/DIV-3.md) | generated |
| SV-1 Systems Interface Description | [SV-1.md](views/SV-1.md) | written |
| SV-4 Systems Functionality | [SV-4.md](views/SV-4.md) | written |
| SV-7 Systems Measures | [SV-7.md](views/SV-7.md) | written |
| SV-8 Systems Evolution | [SV-8.md](views/SV-8.md) | generated |
| StdV-1 Standards Profile | [StdV-1.md](views/StdV-1.md) | written |

## 4. Correspondences

| Rule | How it's kept |
| --- | --- |
| Every OV-5b activity is supported by at least one SV-4 function. | The OV-5b table names its SV-4 function. |
| Every DIV-3 table appears in DIV-2. | Both are generated from the same build. |
| Every state in OV-6b is a value the schema enforces. | `test-architecture-views` checks the OV-6b states against the build's validation rules. |
| Every OV-6a rule cites its source. | Generated from the rule tables, which require a source locator. |
| Every SV-1 interface has a test runner. | SV-1 lists the runner for each interface. |

## 5. Known issues

- Formal DoDAF conformance (DM2/PES) isn't attempted (section 2.4).
- Screen-reader acceptance and printed fit are not yet performed (section 1.3).
- The [DevSecOps alignment matrix](DevSecOps_Alignment.md) maps every guidebook activity to Swift's practice. No gaps remain. Developmental cyber testing is partial: there has been no independent penetration test.
- The [threat model](Threat_Model.md) lists six recommendations. Malformed-input testing is done; code signing (threat model section 5) leads the rest.
