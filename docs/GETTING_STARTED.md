# Getting Started

[Documentation Home](README.md) | [Core Concepts](CORE_CONCEPTS.md) | [Feature Guide](FEATURES.md) | [API Reference](API.md)

This guide gets you from an empty project to a generated `.docx` report.

## Requirements

- .NET Framework 4.6.1+, .NET Core 2.0+, or any modern .NET (5-10) runtime — the package targets `netstandard2.0` with a dedicated `net10.0` build.
- A project that can reference NuGet packages.
- Microsoft Word, LibreOffice, or another DOCX reader for visual validation.

TerraFluent.Docx.Reporting does not require Microsoft Word to be installed on the machine that generates documents.

## Install

```powershell
dotnet add package TerraFluent.Docx.Reporting
```

Use the namespace:

```csharp
using TerraFluent.Docx.Reporting;
```

Use `TerraFluent.Docx.Reporting.Infra` when you implement reusable components:

```csharp
using TerraFluent.Docx.Reporting.Infra;
```

## Create Your First Document

```csharp
using TerraFluent.Docx.Reporting;

Document.Create(doc =>
{
    doc.MetadataTitle("Quarterly Report")
       .MetadataAuthor("Northwind Consulting")
       .MetadataSubject("Q4 performance summary")
       .MetadataKeywords("quarterly report, finance, operations")
       .MetadataCreator("TerraFluent.Docx.Reporting");

    doc.Page(page =>
    {
        page.Size(PageSize.A4);
        page.Margin(Unit.Centimetre(2));

        page.Header().Text("Quarterly Report", text => text
            .Bold()
            .FontColor(Colors.Blue.L800)
            .AlignCenter());

        page.Footer().Text(text =>
        {
            text.Span("Page ");
            text.CurrentPageNumber();
            text.Span(" of ");
            text.TotalPages();
            text.AlignCenter().FontSize(9).FontColor(Colors.Grey.L600);
        });

        page.Content().H1("Executive Summary");
        page.Content().Text("Revenue improved across every practice.");
        page.Content().Text("The report was generated without automating Microsoft Word.");
    });
}).PublishDocx("quarterly-report.docx");
```

## Write To A File, Byte Array, Or Stream

```csharp
var document = Document.Create(doc =>
{
    doc.Page(page => page.Content().Text("Portable output options."));
});

document.PublishDocx("output.docx");

byte[] bytes = document.PublishDocx();

using var stream = File.Create("stream-output.docx");
document.PublishDocx(stream);
```

## Add A Second Section

Each `doc.Page(...)` call adds a new section. Use separate sections when page size, orientation, headers, footers, columns, or page numbering should change.

```csharp
Document.Create(doc =>
{
    doc.Page(page =>
    {
        page.Size(PageSize.A4);
        page.Content().H1("Portrait Summary");
    });

    doc.Page(page =>
    {
        page.Size(PageSize.A4).Landscape();
        page.Content().H1("Landscape Appendix");
    });
}).PublishDocx("multi-section.docx");
```

## Next Steps

- Learn the mental model in [Core Concepts](CORE_CONCEPTS.md).
- Copy larger examples from the [Feature Guide](FEATURES.md).
- Browse runnable sample files in [Samples](SAMPLES.md).

