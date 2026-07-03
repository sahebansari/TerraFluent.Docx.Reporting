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
```

`H1` through `H6` don't accept a bookmark name directly. To bookmark a heading, use `Bookmark` with `Style` instead of `H1`-`H6`:

```csharp
page.Content().Bookmark("operations-section", "Operations", text => text.Style("Heading1"));

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

## Barcodes

```csharp
page.Content().Barcode("SKU-00100011");

page.Content().Barcode("ACME-99887766", bc => bc
    .Width(220)
    .Height(50)
    .BarColor(Colors.Blue.L700)
    .AlignCenter()
    .Caption("Figure 1. Product tracking code"));

page.Content().Barcode("Internal-Only", bc => bc.ShowText(false));
```

Barcodes encode text as Code 128 and render as vector bars (not raster images), so they stay crisp at any size. Only ASCII 32-126 (space through `~`) can be encoded; anything else throws `ArgumentException` immediately.

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

Legend placement, axis titles, data labels, and stacked bars:

```csharp
page.Content().Chart(chart => chart
    .Title("Stacked Revenue by Quarter")
    .Width(430)                           // points; default is 432 x 252 (6 x 3.5 in)
    .Height(250)
    .AlignCenter()
    .Legend(ChartLegendPosition.Bottom)   // or .HideLegend()
    .CategoryAxisTitle("Quarter")
    .ValueAxisTitle("Revenue ($M)")
    .DataLabels()
    .Stacked()                            // or .PercentStacked(); bar charts only
    .Series("Product", s => s.Bar("Q1", 4.1).Bar("Q2", 4.3).Color(Colors.Blue.L700))
    .Series("Services", s => s.Bar("Q1", 2.2).Bar("Q2", 2.6).Color(Colors.Orange.L700)));
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

## Auto-Numbered Captions And Table Of Figures

`FigureCaption` (images) and `Caption` (tables) insert Word `SEQ` fields, so Word renumbers captions automatically and can collect them into a table of figures:

```csharp
page.Content().TableOfFigures();                       // lists Figure captions
page.Content().TableOfFigures("List of Tables", "Table");  // lists Table captions

page.Content().Image("chart.png", img => img
    .Width(200)
    .FigureCaption("Revenue growth chart"));           // "Figure 1. Revenue growth chart"

page.Content().Table(t =>
{
    t.Caption("Quarterly results by region");          // "Table 1. Quarterly results by region", above the table
    t.ColumnsDefinition(d => { d.RelativeColumn(1); d.RelativeColumn(1); });
    t.Row(r => { r.Cell().Text("North"); r.Cell().Text("$4.1M"); });
});
```

Figure captions render below the image; table captions render above the table. The static `Caption(...)` image overload remains available for unnumbered captions.

## Restrict Editing

`RestrictEditing` applies Word's "Restrict Editing" protection, optionally guarded by a password:

```csharp
Document.Create(doc =>
{
    doc.RestrictEditing(DocumentProtection.ReadOnly, "secret123");
    doc.Page(page => page.Content().Text("Read-only content."));
});
```

Modes: `ReadOnly`, `CommentsOnly`, `TrackedChangesOnly`, and `FormsOnly`. The password (first 15 characters) is stored as a salted, spun SHA-512 verifier per ISO/IEC 29500. This is a guard rail, not security: the file is not encrypted, and a malicious tool can strip the setting.

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

