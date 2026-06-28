# API Reference

[Documentation Home](README.md) | [Getting Started](GETTING_STARTED.md) | [Core Concepts](CORE_CONCEPTS.md) | [Feature Guide](FEATURES.md) | [Samples](SAMPLES.md) | [Troubleshooting](TROUBLESHOOTING.md)

This page lists the public fluent API exposed by `TerraFluent.Docx.Reporting`.

## Namespace

```csharp
using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;
```

Most users only need `TerraFluent.Docx.Reporting`. Add `TerraFluent.Docx.Reporting.Infra` when implementing reusable components or receiving descriptor interfaces in your own helper methods.

## Document

| API | Purpose |
| --- | --- |
| `Document.Create(Action<IDocumentContainer> configure)` | Builds a document with the fluent descriptor API. |
| `PublishDocx(string filePath)` | Writes a `.docx` file to disk. |
| `PublishDocx()` | Returns the generated `.docx` package as a byte array. |
| `PublishDocx(Stream stream)` | Writes the generated `.docx` package to a stream. |

```csharp
var document = Document.Create(doc =>
{
    doc.MetadataTitle("Quarterly Report");
    doc.Page(page => page.Content().Text("Hello from TerraFluent.Docx.Reporting."));
});

document.PublishDocx("report.docx");
byte[] bytes = document.PublishDocx();
```

## Document Container

`IDocumentContainer` is the root builder received by `Document.Create`.

| API | Purpose |
| --- | --- |
| `Theme(DocumentTheme theme)` | Applies an existing theme object. |
| `Theme(Action<IDocumentThemeDescriptor> configure)` | Configures the document theme inline. |
| `ParagraphStyle(string name, Action<ITextDescriptor> configure)` | Registers a reusable paragraph style. |
| `TableStyle(string name, Action<ITableDescriptor> configure)` | Registers a reusable table style. |
| `Page(Action<IPageDescriptor> configure)` | Adds a document section/page definition. |
| `MetadataTitle`, `MetadataAuthor`, `MetadataSubject`, `MetadataKeywords`, `MetadataCreator` | Sets package metadata. |
| `Compose(IDocument document)` | Composes a reusable document module. |

```csharp
Document.Create(doc =>
{
    doc.MetadataTitle("Board Pack")
       .MetadataAuthor("Finance Team")
       .MetadataCreator("TerraFluent.Docx.Reporting")
       .Theme(theme => theme
           .DefaultFont("Aptos", 10.5f)
           .HeadingColor(Colors.Blue.L800)
           .AccentColor(Colors.Green.L700));

    doc.ParagraphStyle("FinePrint", text => text.FontSize(8).FontColor(Colors.Grey.L600));
});
```

## Theme

| API | Purpose |
| --- | --- |
| `DefaultFont(string family, float size = 11)` | Sets default family and size. |
| `DefaultFontFamily(string family)` | Sets only the default family. |
| `DefaultFontSize(float size)` | Sets only the default size. |
| `DefaultTextColor(string hexColor)` | Sets default body text color. |
| `HeadingColor(string hexColor)` | Sets heading color. |
| `AccentColor(string hexColor)` | Sets accent color. |
| `HyperlinkColor(string hexColor)` | Sets hyperlink color. |
| `TableHeaderBackground(string hexColor)` | Sets default table header fill. |
| `TableAlternateRowBackground(string hexColor)` | Sets default alternating row fill. |
| `TableBorder(float width, string hexColor)` | Sets default table border style. |
| `TableCellPadding(float points)` | Sets uniform table cell padding. |
| `TableCellPadding(float verticalPoints, float horizontalPoints)` | Sets vertical and horizontal cell padding. |
| `TableRowMinHeight(float points)` | Sets default body row minimum height. |
| `TableHeaderRowMinHeight(float points)` | Sets default header row minimum height. |

## Page

| API | Purpose |
| --- | --- |
| `Size(PageSize size)` / `Size(float widthPoints, float heightPoints)` | Sets page size in points. |
| `Landscape()` / `Portrait()` | Changes orientation. |
| `Margin(...)`, `MarginTop`, `MarginRight`, `MarginBottom`, `MarginLeft` | Sets margins in points. |
| `DefaultTextStyle(Action<ITextDescriptor> configure)` | Sets section default text style. |
| `PageNumberStart(int value)` | Starts page numbering for the section. |
| `PageNumberFormat(string format)` | Writes an OOXML number format such as `decimal`, `lowerRoman`, or `upperLetter`. |
| `Background(string hexColor)` | Sets document page background color. Word stores one document-wide background. |
| `Watermark(string text, string hexColor = "D9D9D9", float fontSize = 54)` | Adds a section watermark. |
| `Columns(int count, float spacingPoints = 36, bool separatorLine = false)` | Enables newspaper-style columns. |
| `SingleColumn()` | Returns the section to one column. |
| `Header`, `OddPageHeader`, `FirstPageHeader`, `EvenPageHeader` | Returns header containers. |
| `Content()` | Returns the body content container. |
| `Footer`, `OddPageFooter`, `FirstPageFooter`, `EvenPageFooter` | Returns footer containers. |

```csharp
doc.Page(page =>
{
    page.Size(PageSize.A4)
        .Margin(Unit.Centimetre(2))
        .PageNumberStart(1)
        .PageNumberFormat("decimal")
        .Watermark("DRAFT", Colors.Grey.L300, 72);

    page.Header().Text("Confidential", t => t.AlignCenter().FontSize(9));
    page.Footer().Text(t =>
    {
        t.Span("Page ");
        t.CurrentPageNumber();
        t.Span(" of ");
        t.TotalPages();
        t.AlignCenter();
    });
});
```

## Container

`IContainer` is used for page body content, headers, footers, columns, row items, and table cells.

| API | Purpose |
| --- | --- |
| `H1` through `H6` | Adds Word heading paragraphs. |
| `Column(Action<IColumnDescriptor>)` | Adds a vertical stack of content. |
| `Row(Action<IRowDescriptor>)` | Adds a horizontal layout row. |
| `Text(string, Action<ITextDescriptor>? configure = null)` | Adds a paragraph. |
| `Text(Action<ITextDescriptor> configure)` | Adds a rich paragraph with multiple runs. |
| `Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null)` | Adds a hyperlink paragraph. |
| `Bookmark(string name)` | Adds an invisible bookmark anchor. |
| `Bookmark(string name, string text, Action<ITextDescriptor>? configure = null)` | Adds a visible bookmarked paragraph. |
| `TableOfContents(string title = "Contents", int minLevel = 1, int maxLevel = 3)` | Adds a Word TOC field. |
| `BulletList` / `NumberedList` | Adds list content. |
| `Table` | Adds a table. |
| `Chart` | Adds a chart. |
| `Image` overloads | Adds an image from path, bytes, or stream. |
| `Line()` | Adds a horizontal rule. |
| `PageBreak()` | Adds a page break. |
| `Component(IComponent component)` | Composes reusable content. |

## Text

| API | Purpose |
| --- | --- |
| `Bold`, `Italic`, `Underline`, `Strikethrough` | Run formatting. |
| `Superscript`, `Subscript` | Vertical text positioning. |
| `FontSize`, `FontColor`, `FontFamily`, `Highlight` | Font styling. |
| `Style(string name)` | Applies a registered paragraph style. |
| `AlignLeft`, `AlignCenter`, `AlignRight`, `Justify` | Paragraph alignment. |
| `LineHeight`, `SpacingBefore`, `SpacingAfter` | Paragraph spacing. |
| `LeftIndent`, `RightIndent`, `FirstLineIndent`, `HangingIndent` | Indentation. |
| `KeepWithNext`, `KeepLinesTogether`, `PageBreakBefore` | Pagination controls. |
| `Shading`, `Border`, `BorderTop`, `BorderRight`, `BorderBottom`, `BorderLeft` | Paragraph background and borders. |
| `TabStop`, `Tab` | Tab layout. |
| `Span` | Adds formatted text runs. |
| `Hyperlink` | Adds inline hyperlink runs. |
| `CrossReference` | Adds an internal reference to a bookmark. |
| `Footnote`, `Endnote` | Adds notes. |
| `CurrentPageNumber`, `TotalPages` | Adds page number fields. |

```csharp
page.Content().Text(t =>
{
    t.Span("Revenue ").Bold();
    t.Span("increased 12%").FontColor(Colors.Green.L700);
    t.Span(" year over year.");
    t.SpacingAfter(8).KeepLinesTogether();
});
```

## Lists

| API | Purpose |
| --- | --- |
| `Marker(string marker, int level = 0, string? fontFamily = null)` | Sets a marker for a list level. |
| `Item(string text, Action<ITextDescriptor>? configure = null)` | Adds a level-0 item. |
| `Item(string text, int level, Action<ITextDescriptor>? configure = null)` | Adds an item at a specific level. |
| `Item(Action<ITextDescriptor> configure)` | Adds a rich level-0 item. |
| `Item(int level, Action<ITextDescriptor> configure)` | Adds a rich item at a specific level. |

## Tables

| API | Purpose |
| --- | --- |
| `Style(string name)` | Applies a registered table style. |
| `Width(float points)` / `WidthPercent(float percent)` | Sets table width. |
| `AlignLeft`, `AlignCenter`, `AlignRight` | Sets table alignment. |
| `ColumnsDefinition(Action<ITableColumnsDefinition>)` | Sets relative or fixed columns. |
| `HeaderRow(Action<ITableRowDescriptor>)` | Adds a repeating-style header row. |
| `Row(Action<ITableRowDescriptor>)` | Adds a body row. |
| `Border`, `CellPadding`, `HeaderBackground`, `AlternateRowBackground` | Sets table formatting. |
| `RowMinHeight`, `HeaderRowMinHeight` | Sets row heights. |

Column APIs:

| API | Purpose |
| --- | --- |
| `RelativeColumn(float size = 1)` | Adds proportional width column. |
| `ConstantColumn(float widthPoints)` | Adds fixed width column. |

Row APIs:

| API | Purpose |
| --- | --- |
| `KeepTogether(bool value = true)` | Keeps a row together. |
| `Cell(int columnSpan = 1)` | Adds a cell; table cells are containers. |

Cell APIs:

| API | Purpose |
| --- | --- |
| `ColumnSpan(int columns)` | Spans multiple columns. |
| `VerticalMergeStart` / `VerticalMergeContinue` | Creates merged vertical cells. |
| `Background`, `Border`, `BorderTop`, `BorderRight`, `BorderBottom`, `BorderLeft` | Cell styling. |
| `Padding(...)` | Cell padding. |
| `VerticalAlignTop`, `VerticalAlignMiddle`, `VerticalAlignBottom` | Vertical alignment. |
| `TextDirectionLeftToRight`, `TextDirectionTopToBottom`, `TextDirectionBottomToTop` | Text direction. |

## Images

| API | Purpose |
| --- | --- |
| `Width`, `Height`, `MaxWidth` | Controls size in points. |
| `AltText` | Adds accessibility text. |
| `Caption` | Adds a caption paragraph. |
| `AlignLeft`, `AlignCenter`, `AlignRight` | Aligns inline image paragraphs. |
| `WrapInline`, `WrapSquare`, `WrapTight`, `WrapTopBottom` | Text wrapping. |
| `BehindText`, `InFrontOfText` | Floating layer. |
| `FloatLeft`, `FloatRight`, `FloatCenter` | Floating alignment. |
| `Position`, `PositionFromPage` | Absolute positioning. |
| `Margin(...)` | Wrap margin. |
| `Border`, `Rounded`, `Crop` | Visual styling. |

## Charts

| API | Purpose |
| --- | --- |
| `Title(string title)` | Sets chart title. |
| `Series(Action<ISeriesDescriptor> configure)` | Adds an unnamed series. |
| `Series(string name, Action<ISeriesDescriptor> configure)` | Adds a named series for legend output. |

Series APIs:

| API | Purpose |
| --- | --- |
| `Bar`, `Line`, `Pie`, `Doughnut` | Adds a data point of that chart kind. |
| `Color(string hexColor)` | Sets the series color. |

All series in a chart must use the same chart kind. Pie and doughnut charts support one series.

## Layout Helpers

Row APIs:

| API | Purpose |
| --- | --- |
| `Spacing(float points)` | Sets spacing between row items. |
| `RelativeItem(float size = 1)` | Adds a proportional width item container. |
| `AutoItem()` | Adds an auto-sized item container. |
| `ConstantItem(float widthPoints)` | Adds a fixed-width item container. |

Column APIs:

| API | Purpose |
| --- | --- |
| `Spacing(float points)` | Sets spacing between column items. |
| `Item()` | Adds an item container. |

## Templates

| API | Purpose |
| --- | --- |
| `DocxTemplate.Open(string templatePath)` | Opens a `.docx` template. |
| `Replace(string placeholder, string value)` | Replaces plain text placeholders. |
| `ReplaceContentControl(string tagOrAlias, string value)` | Replaces matching content controls. |
| `SaveAs(string outputPath)` | Writes the result to disk. |
| `Save()` | Returns the result as a byte array. |

## Utilities

### Unit

| API | Purpose |
| --- | --- |
| `Unit.Centimetre(float value)` | Converts centimetres to points. |
| `Unit.Millimetre(float value)` | Converts millimetres to points. |
| `Unit.Inch(float value)` | Converts inches to points. |
| `Unit.Point(float value)` | Returns points unchanged. |

### PageSize

Built-in sizes: `A0`, `A1`, `A2`, `A3`, `A4`, `A5`, `A6`, `Letter`, `Legal`, `Ledger`, `Tabloid`.

Use `new PageSize(widthPoints, heightPoints)` for custom sizes and `PageSize.A4.Landscape()` for a landscape value.

### Colors

`Colors` provides common hex values under `Black`, `White`, `Grey`, `Blue`, `Red`, `Green`, and `Orange`. You can also pass any six-character hex color string, such as `"1F4E79"`.

### Enums

| Enum | Values |
| --- | --- |
| `HighlightColor` | `Yellow`, `Green`, `Cyan`, `Magenta`, `Blue`, `Red`, `DarkBlue`, `DarkCyan`, `DarkGreen`, `DarkMagenta`, `DarkRed`, `DarkYellow`, `DarkGray`, `LightGray`, `Black`, `White`, `None` |
| `TabStopAlignment` | `Left`, `Center`, `Right`, `Decimal` |

