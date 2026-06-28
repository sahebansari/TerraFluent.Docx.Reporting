# Feature Guide

[Documentation Home](README.md) | [Getting Started](GETTING_STARTED.md) | [Core Concepts](CORE_CONCEPTS.md) | [API Reference](API.md) | [Samples](SAMPLES.md)

This guide shows practical examples for the main document features.

## Document Metadata

```csharp
Document.Create(doc =>
{
    doc.MetadataTitle("Annual Report")
       .MetadataAuthor("Contoso Finance")
       .MetadataSubject("FY2026 performance")
       .MetadataKeywords("annual report, finance, operations")
       .MetadataCreator("TerraFluent.Docx.Reporting");
});
```

## Headers, Footers, And Page Numbers

```csharp
doc.Page(page =>
{
    page.FirstPageHeader().Text("Annual Report", t => t.Bold().AlignCenter());
    page.EvenPageHeader().Text("Contoso", t => t.AlignLeft().FontSize(9));
    page.OddPageHeader().Text("FY2026", t => t.AlignRight().FontSize(9));

    page.Footer().Text(t =>
    {
        t.Span("Page ");
        t.CurrentPageNumber();
        t.Span(" of ");
        t.TotalPages();
        t.AlignCenter().FontSize(9);
    });
});
```

## Page Layout

```csharp
doc.Page(page =>
{
    page.Size(PageSize.A4)
        .Margin(Unit.Centimetre(2))
        .Background("FAFAFA")
        .Watermark("DRAFT", Colors.Grey.L300, 72);

    page.Content().Text("The page background is document-wide in Word.");
});

doc.Page(page =>
{
    page.Size(PageSize.A4).Landscape();
    page.Columns(2, spacingPoints: 24, separatorLine: true);
    page.Content().H1("Two Column Appendix");
    page.Content().Text("Content flows from the first column to the second.");
});
```

## Text And Rich Runs

```csharp
page.Content().Text(text =>
{
    text.Span("Revenue ").Bold();
    text.Span("increased 12%").FontColor(Colors.Green.L700).Bold();
    text.Span(" year over year.");
    text.SpacingAfter(8).KeepLinesTogether();
});

page.Content().Text("Important policy note.", text => text
    .Shading(Colors.Orange.L100)
    .BorderLeft(4, Colors.Orange.L700)
    .LeftIndent(14)
    .RightIndent(14)
    .SpacingBefore(6)
    .SpacingAfter(6));
```

## Links, Bookmarks, Cross References, And Notes

```csharp
page.Content().Bookmark("revenue-section", "Revenue", text => text.Bold());

page.Content().Text(text =>
{
    text.Span("See ");
    text.CrossReference("revenue-section", "the revenue section");
    text.Span(" for details.");
});

page.Content().Text(text =>
{
    text.Span("External reference: ");
    text.Hyperlink("company site", "https://example.com", link => link.FontColor(Colors.Blue.L700));
});

page.Content().Text(text =>
{
    text.Span("Net revenue excludes discontinued products.");
    text.Footnote("A discontinued product is excluded after the final shipment date.");
});
```

## Lists

```csharp
page.Content().BulletList(list =>
{
    list.Marker(">");
    list.Marker("-", level: 1);
    list.Item("Prepare source data");
    list.Item("Validate totals", level: 1, text => text.FontColor(Colors.Green.L700));
    list.Item("Generate final report");
});

page.Content().NumberedList(list =>
{
    list.Item("Open the generated document in Word.");
    list.Item("Confirm there are no repair prompts.");
    list.Item("Export to PDF if needed.");
});
```

## Tables

```csharp
page.Content().Table(table =>
{
    table.WidthPercent(100)
         .CellPadding(4, 6)
         .HeaderBackground(Colors.Blue.L800)
         .AlternateRowBackground(Colors.Grey.L100)
         .Border(0.5f, Colors.Grey.L300);

    table.ColumnsDefinition(cols =>
    {
        cols.RelativeColumn(3);
        cols.ConstantColumn(90);
        cols.ConstantColumn(90);
    });

    table.HeaderRow(row =>
    {
        row.KeepTogether();
        row.Cell().Text("Metric", t => t.Bold().FontColor(Colors.White.Default));
        row.Cell().Text("Q3", t => t.Bold().FontColor(Colors.White.Default).AlignRight());
        row.Cell().Text("Q4", t => t.Bold().FontColor(Colors.White.Default).AlignRight());
    });

    table.Row(row =>
    {
        row.Cell().Text("Revenue");
        row.Cell().Text("$4.1M", t => t.AlignRight());
        row.Cell().Text("$4.8M", t => t.AlignRight().Bold());
    });
});
```

### Spans And Vertical Merges

```csharp
page.Content().Table(table =>
{
    table.ColumnsDefinition(cols =>
    {
        cols.RelativeColumn();
        cols.RelativeColumn();
        cols.RelativeColumn();
    });

    table.Row(row =>
    {
        row.Cell(3).Background(Colors.Blue.L100).Text("Regional Summary", t => t.Bold().AlignCenter());
    });

    table.Row(row =>
    {
        row.Cell().VerticalMergeStart().Text("North");
        row.Cell().Text("Revenue");
        row.Cell().Text("$2.3M");
    });

    table.Row(row =>
    {
        row.Cell().VerticalMergeContinue();
        row.Cell().Text("Margin");
        row.Cell().Text("31%");
    });
});
```

## Images

```csharp
page.Content().Image("logo.png", image => image
    .Width(120)
    .AltText("Company logo")
    .AlignCenter()
    .Caption("Figure 1. Company logo"));

page.Content().Image(File.ReadAllBytes("photo.png"), "photo.png", image => image
    .Width(160)
    .WrapSquare(8)
    .FloatRight(8)
    .Border(1, Colors.Grey.L400)
    .Rounded()
    .Crop(4, 4, 4, 4));
```

## Charts

```csharp
page.Content().Chart(chart => chart
    .Title("Quarterly Revenue")
    .Series("Revenue", series => series
        .Bar("Q1", 4.1)
        .Bar("Q2", 4.3)
        .Bar("Q3", 4.8)
        .Bar("Q4", 5.2)
        .Color(Colors.Green.L700)));
```

Line chart:

```csharp
page.Content().Chart(chart => chart
    .Title("Customer Growth")
    .Series("Customers", series => series
        .Line("Jan", 120)
        .Line("Feb", 132)
        .Line("Mar", 150)
        .Color(Colors.Blue.L700)));
```

Pie chart:

```csharp
page.Content().Chart(chart => chart
    .Title("Revenue Mix")
    .Series(series => series
        .Pie("Consulting", 45)
        .Pie("Support", 30)
        .Pie("Licensing", 25)
        .Color(Colors.Orange.L700)));
```

## Rows And Columns

```csharp
page.Content().Row(row =>
{
    row.Spacing(12);
    row.RelativeItem(2).Text("Main narrative column.");
    row.RelativeItem(1).Text("Sidebar", text => text
        .Shading(Colors.Grey.L100)
        .Border(0.5f, Colors.Grey.L300)
        .LeftIndent(8)
        .RightIndent(8));
});

page.Content().Column(column =>
{
    column.Spacing(6);
    column.Item().H2("Stacked Content");
    column.Item().Text("First block.");
    column.Item().Text("Second block.");
});
```

## Table Of Contents

```csharp
page.Content().TableOfContents("Contents", minLevel: 1, maxLevel: 3);
page.Content().PageBreak();
page.Content().H1("Executive Summary");
page.Content().H2("Revenue");
page.Content().H2("Operations");
```

Word updates TOC fields when the document is opened or when fields are refreshed.

## Templates

Use templates when a `.docx` already exists and you only need to replace placeholders or content controls.

```csharp
DocxTemplate.Open("invoice-template.docx")
    .Replace("{{CustomerName}}", "Ada Lovelace")
    .Replace("{{InvoiceTotal}}", "$1,250.00")
    .ReplaceContentControl("PaymentTerms", "Net 30")
    .SaveAs("invoice-output.docx");
```

To return bytes:

```csharp
byte[] output = DocxTemplate.Open("template.docx")
    .Replace("{{Name}}", "Grace Hopper")
    .Save();
```

