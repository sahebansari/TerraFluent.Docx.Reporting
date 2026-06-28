# Changelog

All notable changes to TerraFluent.Docx.Reporting are documented here.

This project follows semantic versioning. Public releases are distributed as a free MIT-licensed NuGet package.

## Unreleased

- Reworked DOCX template replacement to use structured WordprocessingML text traversal, including placeholders split across Word runs and tagged content controls.
- Added fail-fast validation for public document, page, image, and template APIs.
- Enabled warning-clean builds with warnings treated as errors.
- Expanded CI coverage with Windows/Linux build and test jobs, sample generation, package consumption smoke checks, and LibreOffice conversion smoke checks.

## 1.2.0

- Rebranded the package and public namespaces to `TerraFluent.Docx.Reporting`.
- Added fluent DOCX generation support for reports, invoices, templates, charts, images, headers, footers, notes, watermarks, and page layout helpers.
- Added Open XML validation-focused tests and release packaging checks.
- Multi-targeted the package to `netstandard2.0` and `net10.0`, so it now runs on .NET Framework 4.6.1+, .NET Core 2.0+, and every modern .NET release, in addition to the latest runtime.
