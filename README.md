![TerraFluent.Docx.Reporting](https://raw.githubusercontent.com/sahebansari/TerraFluent.Docx.Reporting/master/src/TerraFluent.Docx.Reporting/icon.png)

# TerraFluent.Docx.Reporting

[![NuGet](https://img.shields.io/nuget/v/TerraFluent.Docx.Reporting.svg)](https://www.nuget.org/packages/TerraFluent.Docx.Reporting)
[![NuGet downloads](https://img.shields.io/nuget/dt/TerraFluent.Docx.Reporting.svg)](https://www.nuget.org/packages/TerraFluent.Docx.Reporting)
[![CI](https://github.com/sahebansari/TerraFluent.Docx.Reporting/actions/workflows/ci.yml/badge.svg)](https://github.com/sahebansari/TerraFluent.Docx.Reporting/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/LICENSE.txt)

TerraFluent.Docx.Reporting is a fluent .NET library for generating professional `.docx` files without automating Microsoft Word. It writes Open XML packages directly and is designed for reports, invoices, proposals, policy documents, and company-branded templates.

📚 **Full documentation, guides, and samples:** 🌐 [https://terrafluent.dev/docx/](https://terrafluent.dev/docx/)

> **New in 1.3.0:** [Code 128 barcodes](#barcodes) rendered as crisp vector shapes, auto-numbered figure/table captions with a table of figures, document protection (Word's "Restrict Editing", with password), and richer charts (legend placement, axis titles, data labels, stacked bars, sizing and alignment). See the [changelog](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/CHANGELOG.md) for details.

## Install

```powershell
dotnet add package TerraFluent.Docx.Reporting
```

Targets `netstandard2.0`, so it runs on .NET Framework 4.6.1+, .NET Core 2.0+, and every modern .NET release, with a dedicated `net10.0` build for the latest runtime.

## Quick Start

```csharp
using TerraFluent.Docx.Reporting;

Document.Create(doc =>
{
    doc.MetadataTitle("Quarterly Report")
       .MetadataAuthor("Northwind Consulting");

    doc.Page(page =>
    {
        page.Size(PageSize.A4);
        page.Margin(Unit.Centimetre(2));
        page.Header().Text("Quarterly Report", t => t.Bold().AlignCenter());
        page.Footer().Text(t =>
        {
            t.Span("Page ");
            t.CurrentPageNumber();
            t.Span(" of ");
            t.TotalPages();
            t.AlignCenter();
        });

        page.Content().H1("Executive Summary");
        page.Content().Text("Revenue improved across every practice.");
    });
}).PublishDocx("quarterly-report.docx");
```

## Feature Examples

### Styled Tables

```csharp
page.Content().Table(table =>
{
    table.CellPadding(5, 7)
         .HeaderBackground(Colors.Blue.L700)
         .AlternateRowBackground(Colors.Grey.L100)
         .HeaderRowMinHeight(24)
         .Border(0.75f, Colors.Grey.L300);

    table.ColumnsDefinition(cols =>
    {
        cols.RelativeColumn(3);
        cols.RelativeColumn(1);
    });

    table.HeaderRow(row =>
    {
        row.Cell().Text("Item", t => t.Bold().FontColor(Colors.White.Default));
        row.Cell().Text("Amount", t => t.Bold().FontColor(Colors.White.Default).AlignRight());
    });

    table.Row(row =>
    {
        row.Cell().Text("Implementation");
        row.Cell().Text("$4,500", t => t.AlignRight());
    });
});
```

### Images

```csharp
page.Content().Image("logo.png", img => img
    .Width(120)
    .AltText("Company logo")
    .AlignCenter()
    .Caption("Figure 1. Company logo"));

page.Content().Image(bytes, "photo.png", img => img
    .Width(140)
    .WrapSquare(8)
    .FloatRight(8)
    .Border(1, Colors.Grey.L400)
    .Rounded()
    .Crop(4, 4, 4, 4));
```

### Barcodes

```csharp
page.Content().Barcode("SKU-00100011");

page.Content().Barcode("ACME-99887766", bc => bc
    .Width(220)
    .Height(50)
    .BarColor(Colors.Blue.L700)
    .AlignCenter()
    .Caption("Figure 1. Product tracking code"));
```

### Paragraph Callouts And Page Columns

```csharp
page.Content().Text("Important policy note.", t => t
    .Shading(Colors.Orange.L100)
    .BorderLeft(4, Colors.Orange.L700)
    .LeftIndent(14)
    .KeepLinesTogether());

doc.Page(page =>
{
    page.Columns(2, spacingPoints: 24, separatorLine: true);
    page.Content().Text("Text flows from the first column into the second.");
});
```

### Templates

```csharp
DocxTemplate.Open("template.docx")
    .Replace("CustomerName", "Ada Lovelace")
    .ReplaceContentControl("InvoiceTotal", "$1,250.00")
    .SaveAs("output.docx");
```

## Samples

Run the sample project:

```powershell
dotnet run --project samples\TerraFluent.Docx.Reporting.Sample\TerraFluent.Docx.Reporting.Sample.csproj
```

Generated documents are written to `Desktop\SampleDocs`, including invoices, an annual report, template replacement output, and a layout feature showcase.

## Documentation

- [Documentation home and main menu](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/README.md)
- [Getting started](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/GETTING_STARTED.md)
- [Core concepts](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/CORE_CONCEPTS.md)
- [Feature guide with examples](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/FEATURES.md)
- [API reference](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/API.md)
- [Samples](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/SAMPLES.md)
- [Troubleshooting](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/TROUBLESHOOTING.md)
- [Release and compatibility checklist](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/docs/RELEASE.md)
- [Changelog](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/CHANGELOG.md)

## Validation

The test suite validates generated packages with the Open XML SDK:

```powershell
dotnet test TerraFluent.Docx.Reporting.sln
dotnet pack src\TerraFluent.Docx.Reporting\TerraFluent.Docx.Reporting.csproj -c Release
```

For application compatibility smoke checks, open the generated sample documents in Microsoft Word and LibreOffice before publishing a release.

## License

TerraFluent.Docx.Reporting is a free public library released under the [MIT License](https://github.com/sahebansari/TerraFluent.Docx.Reporting/blob/master/LICENSE.txt).

NuGet packages are built with MIT license metadata, README content, XML documentation, and symbols.
