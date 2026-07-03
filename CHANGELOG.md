# Changelog

All notable changes to TerraFluent.Docx.Reporting are documented here.

This project follows semantic versioning. Public releases are distributed as a free MIT-licensed NuGet package.

## 1.3.0

- Added `IContainer.Barcode(...)` for Code 128 barcodes, rendered as vector DrawingML shapes (no raster image or media part), usable anywhere an image can go: page body, headers, footers, columns, rows, and table cells. Supports sizing, bar color, alignment, an optional human-readable text line, and captions.
- Added a package icon to the NuGet package.
- Added auto-numbered captions: `IImageDescriptor.FigureCaption(...)` ("Figure 1.") and `ITableDescriptor.Caption(...)` ("Table 1.") emit Word `SEQ` fields that Word renumbers on field update, plus `IContainer.TableOfFigures(...)` to list them.
- Added `IDocumentContainer.RestrictEditing(...)` for Word's "Restrict Editing" protection (read-only, comments-only, tracked-changes-only, forms-only), with optional password stored as an ISO/IEC 29500 salted, spun SHA-512 verifier. Verified against Word itself, including a hex-padding case that the widely-copied MSDN reference implementation gets wrong.
- Added chart options: `Legend(position)`/`HideLegend()`, `CategoryAxisTitle`/`ValueAxisTitle`, `DataLabels()`, and `Stacked()`/`PercentStacked()` for multi-series bar charts.
- Added chart sizing and alignment: `Width(points)`/`Height(points)` (default remains 6 x 3.5 inches) and `AlignLeft`/`AlignCenter`/`AlignRight`, so charts can be fitted to narrow containers such as row columns instead of overflowing at a fixed size.
- Fixed: text containing XML 1.0-illegal control characters (e.g. a stray `\x02`) is now stripped instead of being written raw into `document.xml`, which Word previously reported as a corrupt/repairable file.
- Fixed: `.gif` images added without an explicit `Width`/`Height` now use their real pixel dimensions instead of silently falling back to a fixed placeholder box.

## 1.2.1

- Reworked DOCX template replacement to use structured WordprocessingML text traversal, including placeholders split across Word runs and tagged content controls.
- Added fail-fast validation for public document, page, image, and template APIs.
- Enabled warning-clean builds with warnings treated as errors.
- Expanded CI coverage with Windows/Linux build and test jobs, sample generation, package consumption smoke checks, and LibreOffice conversion smoke checks.
- Completed XML documentation comments across the public fluent API and removed the missing-documentation warning suppression, so the package's generated `.xml` doc file now covers the full public surface.
- Fixed documentation drift in the release checklist and added template-replacement and validation-exception guidance to the troubleshooting guide.

## 1.2.0

- Rebranded the package and public namespaces to `TerraFluent.Docx.Reporting`.
- Added fluent DOCX generation support for reports, invoices, templates, charts, images, headers, footers, notes, watermarks, and page layout helpers.
- Added Open XML validation-focused tests and release packaging checks.
- Multi-targeted the package to `netstandard2.0` and `net10.0`, so it now runs on .NET Framework 4.6.1+, .NET Core 2.0+, and every modern .NET release, in addition to the latest runtime.
