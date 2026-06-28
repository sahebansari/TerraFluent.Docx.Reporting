# Security Policy

TerraFluent.Docx.Reporting generates `.docx` packages and may process user-provided images, templates, text, and metadata. Please report security issues privately so they can be evaluated before public disclosure.

## Supported Versions

Security fixes are considered for the latest released package version and the current `main` branch. Older versions may receive fixes when the issue is severe and the fix can be applied safely.

## Reporting A Vulnerability

Please do not open a public GitHub issue for a suspected vulnerability.

Use GitHub's private vulnerability reporting or create a private security advisory for this repository when available. If private reporting is not enabled, contact the maintainers through the repository owner's GitHub profile or NuGet package contact information, and share only the minimum detail needed to establish contact.

Include:

- Affected TerraFluent.Docx.Reporting version or commit.
- A concise description of the issue.
- Reproduction steps or proof-of-concept code.
- Impact assessment, including whether the issue requires untrusted input.
- Any suggested mitigation or patch, if available.

## Response Expectations

Maintainers will try to acknowledge reports within 7 days. After triage, the expected next steps are:

- Confirm whether the report is a valid security issue.
- Identify affected versions and realistic exploit conditions.
- Prepare a fix, mitigation, or advisory.
- Coordinate disclosure timing with the reporter when practical.

## Scope

Examples of in-scope issues include:

- Unsafe handling of malicious templates, images, or document metadata.
- Generated output that can trigger unsafe behavior in supported document viewers.
- Path traversal, arbitrary file overwrite, or unexpected file access caused by library APIs.
- Dependency vulnerabilities that affect normal library use.

Examples usually out of scope:

- Issues that require already-trusted code execution by the application using the library.
- Denial-of-service cases involving intentionally huge trusted inputs without a practical mitigation.
- Vulnerabilities in Microsoft Word, LibreOffice, .NET, or the Open XML SDK that are not caused by TerraFluent.Docx.Reporting behavior.

## Safe Handling Guidance

Applications using TerraFluent.Docx.Reporting should validate untrusted input before adding it to generated documents, avoid writing output to attacker-controlled paths, and keep the .NET runtime and package dependencies current.
