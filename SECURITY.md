# Security policy

Swift NAVFIT stores personnel data, including SSNs, in a local workspace file.
Please report security problems privately so they can be fixed before they are
widely known.

## Reporting a vulnerability

Use GitHub's private reporting: open the repository's **Security** tab and
choose **Report a vulnerability**. Only the project owner can see the report.

Please include:

- the Swift NAVFIT version (shown in **Settings & Help**) and your Access
  version;
- what an attacker could do, and the steps or a sample file that shows it.

**Never include real personnel data.** Don't attach a populated workspace,
NAVFIT98A database, report PDF or export. Reproduce the problem with the Demo
workspace or made-up values.

Don't report security problems in public issues.

## What to expect

This is a personal project maintained in spare time. You should get an
acknowledgement within a week. Fixes go into the next release, and its release
notes will credit you unless you ask otherwise.

## Supported versions

Only the latest release on the
[releases page](https://github.com/MattMer/swift-navfit/releases/latest) gets
security fixes. Releases tagged `-dev` are test builds.

## Scope

Swift NAVFIT relies on the Windows permissions of the folder that holds the
workspace; it adds no password or encryption of its own, and contributor names
are attribution, not logins. Those are documented design decisions, not
vulnerabilities. See the [threat model](docs/architecture/Threat_Model.md) and
the [technical README](TECHNICAL_README.md#data-handling-and-security) for what
Swift does and doesn't protect against.
