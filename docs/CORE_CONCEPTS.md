# Core Concepts

[Documentation Home](README.md) | [Getting Started](GETTING_STARTED.md) | [Feature Guide](FEATURES.md) | [API Reference](API.md)

TerraFluent.Docx.Reporting uses a fluent builder model. You compose a document by configuring descriptors: document, page, container, table, row, cell, image, chart, and text descriptors.

## Document Flow

```csharp
Document.Create(doc =>
{
    doc.Theme(theme => { });
    doc.ParagraphStyle("Name", text => { });
    doc.TableStyle("Name", table => { });

    doc.Page(page =>
    {
        page.Header().Text("Header");
        page.Content().H1("Title");
        page.Footer().Text("Footer");
    });
});
```

The document owns metadata, themes, reusable styles, and one or more pages. Pages own layout settings and containers. Containers own content.

## Units

The layout unit is points. One inch is 72 points.

```csharp
page.Margin(Unit.Centimetre(2));
page.Margin(Unit.Inch(0.75f));
page.Content().Image("logo.png", img => img.Width(Unit.Millimetre(35)));
```

Use `Unit.Point(value)` when you already have point values.

## Colors

Colors are six-character hex RGB strings without `#`.

```csharp
text.FontColor("1F4E79");
table.Border(0.75f, Colors.Grey.L300);
page.Watermark("DRAFT", Colors.Grey.L300, 72);
```

Built-in palettes include `Black`, `White`, `Grey`, `Blue`, `Red`, `Green`, and `Orange`.

## Themes

Themes configure defaults for a document.

```csharp
doc.Theme(theme => theme
    .DefaultFont("Aptos", 10.5f)
    .DefaultTextColor(Colors.Grey.L900)
    .HeadingColor(Colors.Blue.L800)
    .AccentColor(Colors.Green.L700)
    .HyperlinkColor(Colors.Blue.L700)
    .TableHeaderBackground(Colors.Blue.L800)
    .TableAlternateRowBackground(Colors.Grey.L100)
    .TableBorder(0.5f, Colors.Grey.L300)
    .TableCellPadding(4, 6));
```

## Custom Styles

Register named styles once and apply them later.

```csharp
doc.ParagraphStyle("Callout", text => text
    .Shading(Colors.Blue.L100)
    .BorderLeft(4, Colors.Blue.L700)
    .LeftIndent(12)
    .RightIndent(12)
    .SpacingBefore(6)
    .SpacingAfter(6));

doc.TableStyle("FinancialTable", table => table
    .WidthPercent(100)
    .HeaderBackground(Colors.Blue.L800)
    .AlternateRowBackground(Colors.Grey.L100)
    .CellPadding(4, 6)
    .Border(0.5f, Colors.Grey.L300));

doc.Page(page =>
{
    page.Content().Text("Important board note.", t => t.Style("Callout"));
    page.Content().Table(table => table.Style("FinancialTable"));
});
```

## Containers

Containers are reusable content surfaces. You will see them in:

- `page.Header()`, `page.Content()`, and `page.Footer()`.
- Row items created with `row.RelativeItem()`, `row.AutoItem()`, and `row.ConstantItem(...)`.
- Column items created with `column.Item()`.
- Table cells created with `row.Cell()`.

Because table cells are containers, you can nest headings, text, lists, images, and tables inside cells.

## Sections And Page Settings

Every `doc.Page(...)` call creates a section. Use sections to change orientation, margins, headers, footers, columns, watermarks, and page numbering.

```csharp
doc.Page(page =>
{
    page.Size(PageSize.A4).Portrait();
    page.Content().Text("Normal section.");
});

doc.Page(page =>
{
    page.Size(PageSize.A4).Landscape();
    page.Columns(2, spacingPoints: 24, separatorLine: true);
    page.Content().Text("Landscape two-column appendix.");
});
```

## Reusable Components

Implement `IComponent` for reusable blocks.

```csharp
using TerraFluent.Docx.Reporting.Infra;

public sealed class StatusBanner : IComponent
{
    private readonly string _message;

    public StatusBanner(string message)
    {
        _message = message;
    }

    public void Compose(IContainer container)
    {
        container.Text(_message, text => text
            .Bold()
            .Shading(Colors.Green.L100)
            .BorderLeft(4, Colors.Green.L700)
            .LeftIndent(12)
            .SpacingAfter(8));
    }
}

page.Content().Component(new StatusBanner("All systems operational."));
```

## Validation Mindset

TerraFluent.Docx.Reporting writes Open XML directly. The test suite validates generated packages with the Open XML SDK, but final report layouts should still be visually checked in Word or LibreOffice before a public release.

