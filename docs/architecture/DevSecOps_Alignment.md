# DevSecOps alignment matrix

This matrix maps Swift NAVFIT's development practice to the **DoD Enterprise
DevSecOps Activities & Tools Guidebook** activity tables. It fills the
guidebook's "mission-specific tailoring" space. A local copy of the guidebook
workbook is in `FORMS_REFERENCES_TEMPLATES/Architecture/`.

It is part of the [architecture description](Swift_NAVFIT_Architecture_Description.md)
(backlog M10, step 3).

**Context for the tailoring:** Swift NAVFIT is a single-file Microsoft Access
application:
- It has no server, cloud tenancy, network service, API or container.
- One maintainer develops it in a git repository, with automated native tests
  run through Access COM.
- Releases are files a command copies onto storage it approves for PII.

Many guidebook activities assume a software factory with deployed services.
Those are marked not applicable, with the reason.

## Status key

| Status | Meaning |
| --- | --- |
| **Done** | Practiced today. The evidence column names the script, file or record that shows it. |
| **Partial** | Practiced in part. The evidence column says what is missing. |
| **Gap** | Not practiced yet, and Swift's maintainers could do it. The backlog item is named. |
| **Command** | Belongs to the command, ISSM or a test authority. Swift can supply inputs but can't perform or accept it. |
| **N/A** | Doesn't apply to a standalone file-based application. The reason is given. |

## Summary

Counted over the 118 assessed rows below (activities repeated across phases count once):

| Status | Activities |
| --- | --- |
| Done | 55 |
| Partial | 29 |
| Gap | 0 |
| Command | 14 |
| N/A | 20 |

**No gaps remain.** The SBOM, threat model and malformed-input tests of every file importer are done ([SBOM](SBOM.md), [threat model](Threat_Model.md), `test-malformed-imports`). Dependency checking is partial (inventory only), and developmental cyber testing is partial (no independent penetration test). Most
**Partial** rows are static security analysis (no mature security scanner exists
for VBA), single-maintainer review, and acceptance on a Navy-managed workstation.

## Activities repeated across phases

The guidebook lists these in most phases. Each is assessed once here.

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Test audit | REQUIRED | Done | Every `develop.ps1` run writes `build/workflow/<run>/verification.json` with each runner, its result, the artifact identity and input hashes. Commit messages record which suites ran and their counts. |
| Mission-based cyber risk assessment | REQUIRED | Command | Owned by the command's ISSM. Swift supplies the architecture description, data model (DIV-3), interfaces (SV-1), [SBOM](SBOM.md) and [threat model](Threat_Model.md). |
| SBOM / software composition analysis | REQUIRED | Done | [SBOM.md](SBOM.md) and CycloneDX [sbom.cdx.json](sbom.cdx.json), generated from a fresh build and the VBA source; `test-architecture-views` fails when they drift. No third-party runtime libraries. |
| Software factory risk continuous monitoring | REQUIRED | N/A | No shared software factory; one repository with a local build and test pipeline. |

## Plan

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Change management planning | REQUIRED | Done | [AGENTS.md](../../AGENTS.md): what may change freely, what needs asking, and what needs an explicit request (releases, publishing, signing) |
| Configuration identification | REQUIRED | Done | `APP_VERSION` / `SCHEMA_VERSION` in `modSwift`; the build manifest fingerprints every input; artifacts carry `.build.json` identity |
| Configuration management planning | REQUIRED | Done | [BUILD_PIPELINE.md](../../planning/BUILD_PIPELINE.md); `scripts/build-manifest.json`; an input-drift guard fails a build whose source changes mid-build |
| Database design | PREFERRED | Done | Additive, versioned schema policy; [DIV-2](views/DIV-2.md) and [DIV-3](views/DIV-3.md) |
| Design review | PREFERRED | Partial | Feature proposals (behavior, screens, schema, tests) are reviewed by the product owner before building. No independent design board. |
| DevSecOps process design | REQUIRED | Done | [Development_Workflow.md](../Development_Workflow.md), [CONTRIBUTING.md](../../CONTRIBUTING.md) |
| Documentation version control | REQUIRED | Done | All documentation is in git with the source |
| IaC deployment | REQUIRED | N/A | No infrastructure to provision |
| Project / release planning | REQUIRED | Done | [Priority backlog](../../planning/PRIORITY_BACKLOG.md); dated pre-releases on GitHub |
| Project team onboarding planning | REQUIRED | Partial | CONTRIBUTING, AGENTS and the development workflow cover setup and rules. No formal onboarding checklist. |
| Risk management | REQUIRED | Partial | Each release lists what is not yet accepted, and the backlog tracks risks. No risk register with likelihood and impact. |
| Software requirement analysis | REQUIRED | Done | User review decisions recorded in the backlog; policy requirements traced to BUPERSINST 1610.10H citations ([OV-6a](views/OV-6a.md)) |
| System design | REQUIRED | Done | This architecture description |
| Test deployment | REQUIRED | Done | Blank and synthetic Demo test builds published as GitHub pre-releases with SHA-256 hashes |
| Test plan | REQUIRED | Done | `tests/runners.json` (every runner, area and tier); CONTRIBUTING maps each kind of change to its checks; [screen-reader procedure](../Accessibility_Screen_Reader_Test.md) |
| Threat modeling | PREFERRED | Done | [Threat_Model.md](Threat_Model.md): STRIDE over assets, trust boundaries and interfaces, with prioritized recommendations |

## Develop

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Application code development | PREFERRED | Done | Source in `src/`; forms as text |
| Code commit | REQUIRED | Done | git; commits and pushes only on the product owner's request |
| Code commit logging | REQUIRED | Done | Commit messages state the behavior change, verification and anything not run |
| Code commit scan | REQUIRED | Partial | `check-vba-source.ps1` runs after every edit (editor hook) and in the build. It checks structure (Option Explicit, ASCII-only, line and continuation limits, duplicates), not security. |
| Code review | PREFERRED | Partial | Changes are AI-assisted with product-owner review. No second human reviewer. |
| Component test | PREFERRED | Done | Per-feature harnesses (`tests/*Harness.bas`) |
| Database component test | REQUIRED | Done | Schema validation rules exercised by harnesses; `test-swift-data-import` |
| Database development | PREFERRED | Done | Schema DDL and validation rules in the build; migration in `modMigration` |
| Database functional test | PREFERRED | Done | Persistence, rollback and recovery runners (for example `test-traits`, `test-recovery`) |
| Documentation | REQUIRED | Done | User guide, PROJECT_STATE, architecture description |
| Dynamic analysis | PREFERRED | Partial | Native runtime tests drive the real application through Access COM. No runtime instrumentation tool. |
| Functional test | REQUIRED | Done | Core and extended runner tiers |
| Infrastructure code development | PREFERRED | N/A | No infrastructure |
| Security code development | REQUIRED | Partial | Practices in the source: read-only DAO for imported files with no execution of their code; SSNs never in filenames; no network calls; verified commits; confirmations before destructive actions. No written secure-coding standard. |
| Service functional test | PREFERRED | N/A | No services |
| Static analysis | REQUIRED | Partial | VBA compile check and `check-vba-source`; the build-plan preflight. No security static analysis tool exists for VBA. |
| Static code scan before commit | REQUIRED | Partial | Same as the code commit scan |
| Test development | REQUIRED | Done | AGENTS.md: extend the feature's test with the feature; remove a feature's tests with it |
| Unit test | REQUIRED | Done | Harnesses call module functions directly (for example the counseling and sign-off harnesses) |

## Build

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| API security tests | REQUIRED | N/A | No APIs |
| Build | REQUIRED | Done | `scripts/develop.ps1 -Action Build` from the manifest |
| Build configuration control and audit | REQUIRED | Done | Manifest input hashes in each receipt; the input-drift guard; `test-build-plan` (inventory, dependencies, identity) |
| Component test | REQUIRED | Done | As in Develop |
| Dependency vulnerability checking | REQUIRED | Partial | The SBOM lists every run-time dependency; all are Windows or Office components patched through the Navy's Microsoft update channel. The VBScript dependency was removed in 0.49.0. No automated vulnerability feed is checked. |
| Functional test | REQUIRED | Done | As in Develop |
| Integration test | REQUIRED | Done | End-to-end workflow runners (`test-ux-workflows`, NAVFIT exchange, PDF output) |
| Regression test (includes smoke tests) | REQUIRED | Done | `smoke-test`; BuildAndTest every build; ExtendedTest weekly and before every release |
| Release packaging | REQUIRED | Done | Blank and Demo builders; `test-compiled-distribution` checks a compiled ACCDE candidate |
| Software integration test | REQUIRED | Done | Same as the integration test |
| Static application security test and scan | REQUIRED | Partial | As in static analysis |
| Store artifacts | REQUIRED | Done | GitHub releases with SHA-256 hashes; verified releases also on the project Drive folder |
| System test | PREFERRED | Done | `test-ux-demo` on a populated Demo; startup under normal Office security and with content enabled |

## Test

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| API security tests | REQUIRED | N/A | No APIs |
| Compliance scan | REQUIRED | Partial | Section 508 automated checks (`test-accessibility`). Workstation STIG compliance belongs to the command. |
| Database security test | PREFERRED | Partial | Validation rules and constraints are tested. PII protection relies on folder permissions (decision C01). No database security scanner. |
| Dynamic application security test (DAST) | PREFERRED | N/A | No web or network application to probe |
| Interactive application security test (IAST) | PREFERRED | N/A | Same reason |
| Manual security test | REQUIRED | Done | `test-malformed-imports` (extended tier): 55 generated malformed and control inputs against the PDF, RUAD CSV and NAVFIT98A importers. Each must be rejected cleanly within 20 seconds, with no rows written, the source unchanged and no temporary files left. |
| Performance test | PREFERRED | Partial | M09 benchmarks are recorded in commit messages. No standing performance runner. |
| Regression test | REQUIRED | Done | As in Build |
| Service security test | REQUIRED | N/A | No services |
| Software integration test | REQUIRED | Done | As in Build |
| Static application security test and scan | REQUIRED | Partial | As in Develop |
| Suitability test | PREFERRED | Partial | Product-owner review rounds on test builds. No representative user group yet. |
| System test | REQUIRED | Done | As in Build |

## Release

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Artifacts replication | PREFERRED | Partial | GitHub releases, plus the Drive folder for verified releases |
| Developmental cyber tests | PREFERRED | Partial | Malformed-input testing of every file importer (`test-malformed-imports`). No penetration testing by an independent cyber team. |
| Operational readiness test | REQUIRED | Partial | `release-check.ps1` gates releases. Printed fit, an Acrobat-free workstation and signed launch are still open. |
| Release go / no-go decision | REQUIRED | Done | Releases happen only on the product owner's explicit request, after `release-check` |
| User story review and demonstration | PREFERRED | Done | Product-owner review rounds with screenshots (for example M09 before and after) |

## Deliver

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Configuration integration testing | REQUIRED | Partial | Startup tested under normal Office security on a development workstation. Not yet on a Navy-managed workstation. |
| Deliver released artifacts | REQUIRED | Done | GitHub releases; README links |
| Delivery results review | REQUIRED | Done | Published asset names, sizes and SHA-256 read back and compared after every upload |
| Operational cyber tests | PREFERRED | Command | Requires the operational environment and a cyber test authority |
| Operations team acceptance | REQUIRED | Command | The command accepts the release for use |

## Deploy

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Artifact deployment | PREFERRED | N/A | Users copy one file to approved storage; nothing is deployed |
| Artifact download | PREFERRED | Done | GitHub release download with published hashes |
| Compliance tests | REQUIRED | Command | Workstation STIG and ATO compliance |
| Create linked clone of VM master image | PREFERRED | N/A | No virtual machines |
| Database artifact deployment | PREFERRED | N/A | The database is the delivered file |
| Database schema installation | PREFERRED | Done | The schema is built into Blank builds; older workspaces upgrade through the verified import |
| Formal qualification tests | AS REQUIRED | Command | Not required for a command tool; would be set by a test authority |
| Human systems integration tests | AS REQUIRED | Partial | The attended screen-reader procedure and accessibility checks; product-owner usability reviews |
| Infrastructure provisioning automation | PREFERRED | N/A | No infrastructure |
| Interoperability tests | REQUIRED | Done | NAVFIT98A import, export and return runners; PDF template output and import runners ([SV-1](views/SV-1.md)) |
| Interoperability certifications | AS REQUIRED | Command | Certification authority |
| Mission-oriented developmental tests | REQUIRED | Partial | The Demo workflows follow real reporting cycles with synthetic data. Not yet run by a command in its own environment. |
| Operational assessments | AS REQUIRED | Command | Test authority |
| Operational test and evaluation | AS REQUIRED | Command | Test authority |
| Performance tests | REQUIRED | Partial | As in Test |
| Post-deployment checkout | PREFERRED | Partial | Startup checks after building; no checkout procedure for a command's copy |
| Post-deployment security scan | REQUIRED | Command | The command's workstation scanning |
| User evaluation / feedback | REQUIRED | Done | Review rounds recorded in the backlog; GitHub issue templates for bug reports and feature requests |

## Operate

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Business operations (resource usage, billing) | REQUIRED | N/A | No cloud resources |
| Capacity management | PREFERRED | N/A | No service capacity |
| Chaos engineering | PREFERRED | N/A | No running service |
| Cyber OT&E | REQUIRED | Command | Test authority |
| Logging | REQUIRED | Done | `AuditEvent`, `OperationJournal`, report revisions and events, and provenance, all in the workspace |
| Performance tests | PREFERRED | Partial | As in Test |
| Persistent cyber operations tests | AS REQUIRED | Command | Operational environment |
| Roll forward / roll back | PREFERRED | Done | Transactional saves with tested late rollback; verified backup and restore; import of every earlier schema |
| Sustainability and chaos testing | PREFERRED | Partial | Failure tests: forced Access shutdown during editing (`test-ux-workflows`) and late-rollback injection in the import and persistence runners. No random failure injection. |
| User evaluation / feedback | REQUIRED | Done | As in Deploy |

## Monitor

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Asset inventory, including SBOMs | REQUIRED | Command | The command inventories its workstations and software; Swift supplies its [SBOM](SBOM.md). |
| Compliance monitoring (COTS) | REQUIRED | Command | STIG monitoring of Office and Access on command workstations |
| Compliance monitoring (resources and services) | REQUIRED | N/A | No cloud resources or services |
| Database monitoring and security auditing | PREFERRED | Partial | The audit history is viewable in the application (history, provenance, operations). No monitoring service. |
| Feedback, including OT&E as needed | PREFERRED | Done | As in Feedback |
| Log analysis and auditing | REQUIRED | Partial | Logs stay in the workspace and can be reviewed per report or operation. No aggregation or alerting. |
| Log auditing | REQUIRED | Partial | Same |
| Runtime application self-protection (RASP) | PREFERRED | N/A | No runtime instrumentation platform for Access |
| System configuration monitoring | PREFERRED | N/A | Nothing deployed to monitor; workstation configuration is the command's |
| System performance monitoring | PREFERRED | N/A | No running service |
| System security monitoring | REQUIRED | Command | The command's endpoint security |
| Test configuration audit | PREFERRED | Done | Receipts record the input hashes and artifact identity that each test ran against |
| User evaluation / feedback | REQUIRED | Done | As in Deploy |

## Feedback

| Activity | Baseline | Status | Evidence or reason |
| --- | --- | --- | --- |
| Revise product backlog | REQUIRED | Done | [Priority backlog](../../planning/PRIORITY_BACKLOG.md), updated after every review round and finding |
| User evaluation / feedback | REQUIRED | Done | As in Deploy |
