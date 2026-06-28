# TerraFluent.Docx.Reporting Documentation

Welcome to the TerraFluent.Docx.Reporting documentation. Use this page as the main menu for the complete library docs.

## Main Menu

| Section | What You Will Find |
| --- | --- |
| [Getting Started](GETTING_STARTED.md) | Installation, requirements, your first document, and output options. |
| [Core Concepts](CORE_CONCEPTS.md) | How the fluent builder works, descriptors, units, colors, themes, styles, and reusable components. |
| [Feature Guide](FEATURES.md) | End-to-end examples for pages, text, tables, images, lists, charts, templates, and document layout. |
| [API Reference](API.md) | Public API tables for every descriptor, helper type, and template API. |
| [Samples](SAMPLES.md) | Links to runnable sample files and what each sample demonstrates. |
| [Troubleshooting](TROUBLESHOOTING.md) | Common issues with Open XML, images, package metadata, Word repair prompts, and layout. |
| [Release And Publishing](RELEASE.md) | Build, test, pack, inspect, and publish the NuGet package. |
| [Changelog](../CHANGELOG.md) | Version history and notable changes. |
| [License](../LICENSE.txt) | MIT license. |

## Quick Links

- [Install from NuGet](GETTING_STARTED.md#install)
- [Create a basic report](GETTING_STARTED.md#create-your-first-document)
- [Add headers and footers](FEATURES.md#headers-footers-and-page-numbers)
- [Build tables](FEATURES.md#tables)
- [Insert images](FEATURES.md#images)
- [Build charts](FEATURES.md#charts)
- [Use Word templates](FEATURES.md#templates)
- [Reusable components](CORE_CONCEPTS.md#reusable-components)
- [Publish to NuGet](RELEASE.md#publish)

## Library At A Glance

TerraFluent.Docx.Reporting is a fluent .NET library for generating `.docx` reports without automating Microsoft Word. It writes Open XML packages directly and supports:

- Document metadata, sections, headers, footers, page numbering, watermarks, page background, and columns.
- Styled text, headings, hyperlinks, bookmarks, cross references, footnotes, endnotes, callouts, and tab stops.
- Tables with headers, widths, fixed and relative columns, borders, cell padding, spans, vertical merges, and nested content.
- Images from files, byte arrays, or streams with sizing, captions, wrapping, floating, borders, rounded corners, and cropping.
- Lists, charts, reusable themes, custom paragraph/table styles, reusable components, and `.docx` template replacement.

## Minimal Example

```csharp
using TerraFluent.Docx.Reporting;

Document.Create(doc =>
{
    doc.MetadataTitle("Quarterly Report")
       .MetadataAuthor("Finance Team");

    doc.Page(page =>
    {
        page.Size(PageSize.A4);
        page.Margin(Unit.Centimetre(2));
        page.Header().Text("Quarterly Report", t => t.Bold().AlignCenter());
        page.Content().H1("Executive Summary");
        page.Content().Text("Revenue increased across every business unit.");
    });
}).PublishDocx("quarterly-report.docx");
```

Next: [Getting Started](GETTING_STARTED.md).

