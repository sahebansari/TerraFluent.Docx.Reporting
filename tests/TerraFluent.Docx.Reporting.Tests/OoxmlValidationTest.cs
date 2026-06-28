using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Tests;

public class OoxmlValidationTest
{
    [Fact]
    public void SimpleDocument_PassesOoxmlValidation()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Column(col =>
                {
                    col.Item().H1("Hello World");
                    col.Item().Text("A simple paragraph.");
                    col.Item().Text(t => t.Span("Bold").Bold().Span(" and ").Span("italic").Italic());
                });
            });
        }).PublishDocx();

        var errors = Validate(bytes);
        Assert.Empty(errors);
    }

    [Fact]
    public void TableDocument_PassesOoxmlValidation()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Column(col =>
                {
                    col.Item().Table(t =>
                    {
                        t.ColumnsDefinition(d => { d.RelativeColumn(2); d.RelativeColumn(1); });
                        t.Border(1f, Colors.Grey.L400);
                        t.HeaderRow(r => { r.Cell().Text("Name", x => x.Bold()); r.Cell().Text("Score", x => x.Bold()); });
                        t.Row(r => { r.Cell().Text("Alice"); r.Cell().Text("98"); });
                    });
                });
            });
        }).PublishDocx();

        var errors = Validate(bytes);
        Assert.Empty(errors);
    }

    [Fact]
    public void HeaderFooter_PassesOoxmlValidation()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Header().Text(t => t.Span("My Header").Bold());
                p.Footer().Text(t =>
                {
                    t.Span("Page ");
                    t.CurrentPageNumber();
                    t.Span(" of ");
                    t.TotalPages();
                    t.AlignCenter();
                });
                p.Content().Column(col =>
                {
                    col.Item().Text("Body content.");
                });
            });
        }).PublishDocx();

        var errors = Validate(bytes);
        Assert.Empty(errors);
    }

    [Fact]
    public void RowLayout_PassesOoxmlValidation()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Column(col =>
                {
                    col.Item().Row(row =>
                    {
                        row.RelativeItem().Text("Left cell.");
                        row.RelativeItem().Text("Right cell.");
                    });
                });
            });
        }).PublishDocx();

        var errors = Validate(bytes);
        Assert.Empty(errors);
    }

    [Fact]
    public void RowLayout_RelativeAndAutoItem_RelativeCellDoesNotClaimFullWidth()
    {
        // Regression: when a RelativeItem and AutoItem share a row, the relative cell
        // was incorrectly assigned 5000 pct (100%), making the layout impossible for Word.
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Column(col =>
                {
                    col.Item().Row(row =>
                    {
                        row.RelativeItem().Text("Title");
                        row.AutoItem().Text(t => t.CurrentPageNumber().AlignRight());
                    });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        // Relative cell must be 4000 pct (80%), not 5000 (100%)
        Assert.DoesNotContain("w:type=\"pct\" w:w=\"5000\"", documentXml);
        Assert.Contains("w:type=\"pct\"", documentXml);
    }

    [Fact]
    public void PlainParagraph_DoesNotEmitEmptyPPrElement()
    {
        // Regression: paragraphs with default alignment/no style were emitting
        // empty <w:pPr></w:pPr> elements which caused Word's recovery dialog.
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Column(col =>
                {
                    col.Item().Text("Plain paragraph with default properties.");
                    col.Item().Text("Another plain paragraph.");
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.DoesNotContain("<w:pPr></w:pPr>", documentXml);
        Assert.DoesNotContain("<w:pPr/>", documentXml);
    }

    [Fact]
    public void CoreProperties_UseInvariantW3cDateTimeSeparators()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        var originalUiCulture = CultureInfo.CurrentUICulture;
        var culture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        culture.DateTimeFormat.TimeSeparator = ".";

        try
        {
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;

            var bytes = Document.Create(c =>
            {
                c.MetadataTitle("Date Test");
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Margin(Unit.Centimetre(2.54f));
                    p.Content().Text("Body content.");
                });
            }).PublishDocx();

            var coreXml = ReadZipEntry(bytes, "docProps/core.xml");
            Assert.Matches("""<dcterms:created xsi:type="dcterms:W3CDTF">\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z</dcterms:created>""", coreXml);
            Assert.Matches("""<dcterms:modified xsi:type="dcterms:W3CDTF">\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z</dcterms:modified>""", coreXml);
            Assert.DoesNotMatch("""T\d{2}\.\d{2}\.\d{2}Z""", coreXml);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUiCulture;
        }
    }

    [Fact]
    public void DefaultTextStyle_AppliesToPageText()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.DefaultTextStyle(t => t.FontFamily("Arial").FontSize(11).FontColor(Colors.Blue.L700).SpacingAfter(4));
                p.Content().H1("Preserves Heading Style");
                p.Content().Text("Styled by default.");
                p.Footer().Text(t => t.CurrentPageNumber());
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var footerXml = ReadZipEntry(bytes, "word/footer1.xml");
        Assert.Contains("""<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>""", documentXml);
        Assert.Contains("""<w:sz w:val="22"/><w:szCs w:val="22"/>""", documentXml);
        Assert.Contains("""<w:color w:val="1976D2"/>""", documentXml);
        Assert.Contains("""<w:spacing w:before="0" w:after="80"/>""", documentXml);
        Assert.Contains("""<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>""", footerXml);
        Assert.Contains("""<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Preserves Heading Style</w:t></w:r>""", documentXml);
    }

    [Fact]
    public void RowSpacing_EmitsTableCellSpacing()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Row(row =>
                {
                    row.Spacing(6);
                    row.RelativeItem().Text("Left");
                    row.RelativeItem().Text("Right");
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:tblCellSpacing w:w="120" w:type="dxa"/>""", documentXml);
    }

    [Fact]
    public void BulletAndNumberedLists_UseNumberingPart()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().BulletList(list =>
                {
                    list.Item("First bullet");
                    list.Item(t => t.Span("Styled bullet", s => s.Bold()));
                });
                p.Content().NumberedList(list =>
                {
                    list.Item("First step");
                    list.Item("Second step");
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var relsXml = ReadZipEntry(bytes, "word/_rels/document.xml.rels");
        var numberingXml = ReadZipEntry(bytes, "word/numbering.xml");
        Assert.Contains("""<w:numId w:val="1"/>""", documentXml);
        Assert.Contains("""<w:numId w:val="2"/>""", documentXml);
        Assert.Contains("numbering.xml", relsXml);
        Assert.Contains("""<w:numFmt w:val="bullet"/>""", numberingXml);
        Assert.Contains("""<w:multiLevelType w:val="hybridMultilevel"/>""", numberingXml);
        Assert.Contains("""<w:lvl w:ilvl="1">""", numberingXml);
        Assert.Contains("""<w:numFmt w:val="decimal"/>""", numberingXml);
    }

    [Fact]
    public void MultipleDefaultNumberedLists_EachGetOwnNumIdAndRestart()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().NumberedList(list =>
                {
                    list.Item("First step");
                    list.Item("Second step");
                });
                p.Content().Text("Some text between the two lists.");
                p.Content().NumberedList(list =>
                {
                    list.Item("First step again");
                    list.Item("Second step again");
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var numberingXml = ReadZipEntry(bytes, "word/numbering.xml");

        // Two independent numbered lists with identical (default) formatting share one
        // abstract numbering definition but must each get their own <w:num> instance,
        // otherwise the second list continues numbering from the first instead of
        // restarting at 1.
        Assert.Equal(1, CountOccurrences(numberingXml, "<w:abstractNum "));
        Assert.Equal(2, CountOccurrences(numberingXml, "<w:num "));
        Assert.Contains("""<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>""", numberingXml);
        Assert.Contains("""<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>""", numberingXml);
    }

    [Fact]
    public void Hyperlinks_WriteExternalRelationshipsInCorrectScopes()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Header().Hyperlink("Header link", "https://example.com/header");
                p.Footer().Text(t => t.Span("Footer ").Hyperlink("link", "https://example.com/footer"));
                p.Content().Text(t =>
                {
                    t.Span("Visit ");
                    t.Hyperlink("TerraFluent.Docx.Reporting", "https://example.com/terra?x=1&y=2");
                    t.Span(" for details.");
                });
                p.Content().Hyperlink("Standalone link", "https://example.com/standalone");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var documentRels = ReadZipEntry(bytes, "word/_rels/document.xml.rels");
        var headerRels = ReadZipEntry(bytes, "word/_rels/header1.xml.rels");
        var footerRels = ReadZipEntry(bytes, "word/_rels/footer1.xml.rels");

        Assert.Contains("<w:hyperlink r:id=\"", documentXml);
        Assert.Contains("Target=\"https://example.com/terra?x=1&amp;y=2\" TargetMode=\"External\"", documentRels);
        Assert.Contains("Target=\"https://example.com/standalone\" TargetMode=\"External\"", documentRels);
        Assert.Contains("Target=\"https://example.com/header\" TargetMode=\"External\"", headerRels);
        Assert.Contains("Target=\"https://example.com/footer\" TargetMode=\"External\"", footerRels);
    }

    [Fact]
    public void ParagraphControls_EmitIndentationAndPaginationProperties()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text("Controlled paragraph.", t => t
                    .LeftIndent(18)
                    .RightIndent(12)
                    .FirstLineIndent(9)
                    .SpacingBefore(6)
                    .SpacingAfter(8)
                    .KeepWithNext()
                    .KeepLinesTogether()
                    .PageBreakBefore());
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("<w:keepNext/><w:keepLines/><w:pageBreakBefore/>", documentXml);
        Assert.Contains("""<w:spacing w:before="120" w:after="160"/>""", documentXml);
        Assert.Contains("""<w:ind w:left="360" w:right="240" w:firstLine="180"/>""", documentXml);
    }

    [Fact]
    public void ParagraphBordersAndShading_EmitCalloutFormatting()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text("Executive callout.", t => t
                    .Shading(Colors.Blue.L100)
                    .Border(1.25f, Colors.Blue.L700, 5)
                    .LeftIndent(10)
                    .RightIndent(10)
                    .SpacingBefore(8)
                    .SpacingAfter(8));
                p.Content().Text("Warning note.", t => t
                    .Shading(Colors.Orange.L100)
                    .BorderLeft(3, Colors.Orange.L700, 4));
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:pBdr><w:top w:val="single" w:sz="10" w:space="5" w:color="1976D2"/>""", documentXml);
        Assert.Contains("""<w:left w:val="single" w:sz="10" w:space="5" w:color="1976D2"/>""", documentXml);
        Assert.Contains("""<w:bottom w:val="single" w:sz="10" w:space="5" w:color="1976D2"/>""", documentXml);
        Assert.Contains("""<w:right w:val="single" w:sz="10" w:space="5" w:color="1976D2"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="BBDEFB"/>""", documentXml);
        Assert.Contains("""<w:pBdr><w:left w:val="single" w:sz="24" w:space="4" w:color="F57C00"/></w:pBdr>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="FFE0B2"/>""", documentXml);
    }

    [Fact]
    public void PageColumns_EmitSectionColumnSettings()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Columns(3, spacingPoints: 18, separatorLine: true);
                p.Content().Text("Column one text.");
                p.Content().Text("Column two text.");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:cols w:num="3" w:space="360" w:sep="1"/>""", documentXml);
    }

    [Fact]
    public void Lists_SupportNestedLevelsAndCustomMarkers()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().BulletList(list =>
                {
                    list.Marker("✓", fontFamily: "Segoe UI Symbol");
                    list.Marker("→", level: 1, fontFamily: "Segoe UI Symbol");
                    list.Item("Top item");
                    list.Item("Nested item", level: 1);
                });
                p.Content().NumberedList(list =>
                {
                    list.Marker("%1)", level: 0);
                    list.Marker("%2)", level: 1);
                    list.Item("First");
                    list.Item("First child", level: 1);
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var numberingXml = ReadZipEntry(bytes, "word/numbering.xml");

        Assert.Contains("""<w:ilvl w:val="1"/>""", documentXml);
        Assert.Contains("""<w:lvlText w:val="✓"/>""", numberingXml);
        Assert.Contains("""<w:lvlText w:val="→"/>""", numberingXml);
        Assert.Contains("""<w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol" w:hint="default"/>""", numberingXml);
        Assert.Contains("""<w:lvlText w:val="%1)"/>""", numberingXml);
        Assert.Contains("""<w:lvlText w:val="%2)"/>""", numberingXml);
    }

    [Fact]
    public void StyledTable_EmitsPaddingShadingAndRowHeights()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Table(t =>
                {
                    t.CellPadding(5, 7)
                        .HeaderBackground(Colors.Blue.L700)
                        .AlternateRowBackground(Colors.Grey.L100)
                        .HeaderRowMinHeight(24)
                        .RowMinHeight(22)
                        .Border(0.75f, Colors.Grey.L300);
                    t.ColumnsDefinition(d => { d.RelativeColumn(2); d.RelativeColumn(1); });
                    t.HeaderRow(r => { r.Cell().Text("Item"); r.Cell().Text("Amount"); });
                    t.Row(r => { r.Cell().Text("A"); r.Cell().Text("$10"); });
                    t.Row(r => { r.Cell().Text("B"); r.Cell().Text("$20"); });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tblCellMar>""", documentXml);
        Assert.Contains("""<w:tblHeader/>""", documentXml);
        Assert.Contains("""<w:trHeight w:val="480" w:hRule="atLeast"/>""", documentXml);
        Assert.Contains("""<w:trHeight w:val="440" w:hRule="atLeast"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="1976D2"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>""", documentXml);
    }

    [Fact]
    public void StyledTableCells_EmitCellOverridesAndColumnSpan()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Table(t =>
                {
                    t.CellPadding(3, 5).Border(0.75f, Colors.Grey.L300);
                    t.ColumnsDefinition(d =>
                    {
                        d.RelativeColumn(1);
                        d.RelativeColumn(1);
                        d.RelativeColumn(1);
                    });
                    t.Row(r =>
                    {
                        r.Cell(2)
                            .Background(Colors.Grey.L100)
                            .Padding(6, 8)
                            .VerticalAlignMiddle()
                            .Text("Subtotal", x => x.AlignRight().Bold());
                        r.Cell()
                            .Background(Colors.Blue.L700)
                            .VerticalAlignBottom()
                            .Text("$100.00", x => x.AlignRight().Bold().FontColor(Colors.White.Default));
                    });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:gridSpan w:val="2"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>""", documentXml);
        Assert.Contains("""<w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>""", documentXml);
        Assert.Contains("""<w:vAlign w:val="center"/>""", documentXml);
        Assert.Contains("""<w:vAlign w:val="bottom"/>""", documentXml);
    }

    [Fact]
    public void Images_SupportLayoutMetadataAndMemorySources()
    {
        var pngBytes = CreateMinimalPng(1, 1);
        using var pngStream = new MemoryStream(pngBytes);

        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Image(pngBytes, "memory.png", img => img
                    .Width(100)
                    .MaxWidth(50)
                    .AltText("Memory image alt text")
                    .AlignCenter()
                    .Caption("Figure 1. Memory-backed image"));
                p.Content().Image(pngStream, "stream.png", img => img
                    .Height(30)
                    .AltText("Stream image alt text")
                    .AlignRight());
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var relsXml = ReadZipEntry(bytes, "word/_rels/document.xml.rels");

        Assert.Contains("""<w:jc w:val="center"/>""", documentXml);
        Assert.Contains("""<w:jc w:val="right"/>""", documentXml);
        Assert.Contains("descr=\"Memory image alt text\"", documentXml);
        Assert.Contains("descr=\"Stream image alt text\"", documentXml);
        Assert.Contains("""<wp:extent cx="635000" cy="635000"/>""", documentXml);
        Assert.Contains("Figure 1. Memory-backed image", documentXml);
        Assert.Contains("Target=\"media/image1.png\"", relsXml);
        Assert.Contains("Target=\"media/image2.png\"", relsXml);
        Assert.NotNull(ReadZipEntry(bytes, "word/media/image1.png"));
        Assert.NotNull(ReadZipEntry(bytes, "word/media/image2.png"));
    }

    [Fact]
    public void Images_SupportFloatingWrappingCroppingBordersAndRoundedCorners()
    {
        var pngBytes = CreateMinimalPng(1, 1);

        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text("Paragraph before floating image.");
                p.Content().Image(pngBytes, "floating.png", img => img
                    .Width(120)
                    .WrapSquare(8)
                    .FloatRight(8)
                    .Position(24, 12)
                    .Border(1.5f, Colors.Blue.L700)
                    .Rounded()
                    .Crop(5, 10, 5, 0)
                    .AltText("Floating wrapped image"));
                p.Content().Text("Paragraph after floating image.");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("<wp:anchor ", documentXml);
        Assert.Contains("distT=\"101600\" distB=\"101600\" distL=\"101600\" distR=\"101600\"", documentXml);
        Assert.Contains("""<wp:positionH relativeFrom="margin"><wp:posOffset>304800</wp:posOffset></wp:positionH>""", documentXml);
        Assert.Contains("""<wp:positionV relativeFrom="paragraph"><wp:posOffset>152400</wp:posOffset></wp:positionV>""", documentXml);
        Assert.Contains("""<wp:wrapSquare wrapText="bothSides"/>""", documentXml);
        Assert.Contains("""<a:srcRect l="5000" t="10000" r="5000"/>""", documentXml);
        Assert.Contains("""<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>""", documentXml);
        Assert.Contains("""<a:ln w="19050"><a:solidFill><a:srgbClr val="1976D2"/></a:solidFill></a:ln>""", documentXml);
        Assert.Contains("descr=\"Floating wrapped image\"", documentXml);
    }

    [Fact]
    public void Images_WrapTopBottom_DefaultsToCenteredHorizontalPosition()
    {
        var pngBytes = CreateMinimalPng(1, 1);

        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Image(pngBytes, "banner.png", img => img
                    .Width(120)
                    .WrapTopBottom());
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>""", documentXml);
        Assert.Contains("""<wp:wrapTopAndBottom/>""", documentXml);
    }

    [Fact]
    public void Images_WrapTopBottom_WithExplicitAlignment_OverridesDefaultCenter()
    {
        var pngBytes = CreateMinimalPng(1, 1);

        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Image(pngBytes, "banner.png", img => img
                    .Width(120)
                    .AlignRight()
                    .WrapTopBottom());
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>""", documentXml);
    }

    [Fact]
    public void Images_WrapSquare_DefaultsToFlushLeftHorizontalPosition()
    {
        var pngBytes = CreateMinimalPng(1, 1);

        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Image(pngBytes, "wrap.png", img => img
                    .Width(120)
                    .WrapSquare());
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>""", documentXml);
    }

    [Fact]
    public void FullSampleWithImage_PassesOoxmlValidation()
    {
        // Minimal valid 1x1 red PNG (IHDR + IDAT + IEND)
        var pngBytes = CreateMinimalPng(1, 1);
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.png");
        File.WriteAllBytes(pngPath, pngBytes);

        try
        {
            var bytes = Document.Create(c =>
            {
                c.MetadataTitle("Test").MetadataAuthor("Test");
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Margin(Unit.Centimetre(2.54f));

                    p.Header().Column(hdr =>
                    {
                        hdr.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Header Title", t => t.Bold());
                            row.AutoItem().Text(t => t.CurrentPageNumber().AlignRight());
                        });
                        hdr.Item().Line();
                    });

                    p.Footer().Column(ftr =>
                    {
                        ftr.Item().Line();
                        ftr.Item().Text(t =>
                        {
                            t.Span("Page ");
                            t.CurrentPageNumber(s => s.FontSize(9));
                            t.Span(" of ");
                            t.TotalPages(s => s.FontSize(9));
                            t.AlignCenter();
                        });
                    });

                    p.Content().Column(col =>
                    {
                        col.Spacing(8);
                        col.Item().H1("Heading 1");
                        col.Item().Image(pngPath, 100);
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Left");
                            row.RelativeItem().Image(pngPath, 50);
                        });
                        col.Item().Table(t =>
                        {
                            t.ColumnsDefinition(d => { d.RelativeColumn(2); d.RelativeColumn(1); });
                            t.Border(1f, Colors.Grey.L400);
                            t.HeaderRow(r => { r.Cell().Text("A", x => x.Bold()); r.Cell().Text("B", x => x.Bold()); });
                            t.Row(r => { r.Cell().Text("1"); r.Cell().Text("2"); });
                        });
                    });
                });
            }).PublishDocx();

            var errors = Validate(bytes);
            Assert.Empty(errors);
        }
        finally
        {
            File.Delete(pngPath);
        }
    }

    [Fact]
    public void HeaderFooterImages_WritePartScopedRelationships()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.png");
        File.WriteAllBytes(pngPath, CreateMinimalPng(1, 1));

        try
        {
            var bytes = Document.Create(c =>
            {
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Margin(Unit.Centimetre(2.54f));
                    p.Header().Image(pngPath, 16);
                    p.Footer().Image(pngPath, 16);
                    p.Content().Image(pngPath, 16);
                });
            }).PublishDocx();

            Assert.Empty(Validate(bytes));

            var documentRels = ReadZipEntry(bytes, "word/_rels/document.xml.rels");
            var headerRels = ReadZipEntry(bytes, "word/_rels/header1.xml.rels");
            var footerRels = ReadZipEntry(bytes, "word/_rels/footer1.xml.rels");

            Assert.Contains("Target=\"media/image1.png\"", documentRels);
            Assert.Contains("Target=\"media/image1.png\"", headerRels);
            Assert.Contains("Target=\"media/image1.png\"", footerRels);
        }
        finally
        {
            File.Delete(pngPath);
        }
    }

    [Fact]
    public void SpanFluentStyling_AppliesToNewSpanWithoutEmptyRuns()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text(t =>
                {
                    t.Span("Bold, ").Bold();
                    t.Span("italic").Italic();
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.DoesNotContain("<w:t xml:space=\"preserve\"></w:t>", documentXml);
        Assert.Contains("<w:b/>", documentXml);
        Assert.Contains("<w:t xml:space=\"preserve\">Bold, </w:t>", documentXml);
        Assert.Contains("<w:i/>", documentXml);
        Assert.Contains("<w:t xml:space=\"preserve\">italic</w:t>", documentXml);
    }

    [Fact]
    public void SuperscriptAndSubscript_EmitVertAlign()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text(t =>
                {
                    t.Span("E=mc");
                    t.Span("2").Superscript();
                    t.Span(" and H");
                    t.Span("2").Subscript();
                    t.Span("O");
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:vertAlign w:val="superscript"/>""", documentXml);
        Assert.Contains("""<w:vertAlign w:val="subscript"/>""", documentXml);
    }

    [Fact]
    public void Highlight_EmitsHighlightColor()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text(t =>
                {
                    t.Span("Highlighted").Highlight();
                    t.Span(" and ");
                    t.Span("custom color").Highlight(HighlightColor.Green);
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:highlight w:val="yellow"/>""", documentXml);
        Assert.Contains("""<w:highlight w:val="green"/>""", documentXml);
    }

    [Fact]
    public void ParagraphLevelFormatting_AfterSpans_AppliesToRunsWithoutOwnOverride()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text(t =>
                {
                    t.Span("Important: ").Bold();
                    t.Span("Review before release.");
                    t.FontFamily("Aptos")
                     .FontSize(11)
                     .FontColor(Colors.Grey.L900);
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");

        // "Important: " keeps its explicit Bold while also inheriting the paragraph-level font settings.
        Assert.Contains(
            """<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:b/><w:color w:val="212121"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">Important: </w:t>""",
            documentXml);

        // "Review before release." has no per-span overrides, so it picks up the paragraph-level font settings.
        Assert.Contains(
            """<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:color w:val="212121"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">Review before release.</w:t>""",
            documentXml);
    }

    [Fact]
    public void Theme_AppliesDocumentDefaultsHeadingColorHyperlinkColorAndTableDefaults()
    {
        var bytes = Document.Create(c =>
        {
            c.Theme(t => t
                .DefaultFont("Aptos", 10)
                .DefaultTextColor(Colors.Grey.L800)
                .HeadingColor(Colors.Green.L700)
                .HyperlinkColor(Colors.Red.L700)
                .TableBorder(0.75f, Colors.Green.L700)
                .TableCellPadding(4, 6)
                .TableHeaderBackground(Colors.Green.L700)
                .TableAlternateRowBackground(Colors.Green.L100)
                .TableHeaderRowMinHeight(24)
                .TableRowMinHeight(20));

            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().H1("Theme heading");
                p.Content().Text(t => t.Hyperlink("Theme link", "https://example.com/theme"));
                p.Content().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn();
                        cols.RelativeColumn();
                    });
                    table.HeaderRow(row => { row.Cell().Text("A"); row.Cell().Text("B"); });
                    table.Row(row => { row.Cell().Text("1"); row.Cell().Text("2"); });
                    table.Row(row => { row.Cell().Text("3"); row.Cell().Text("4"); });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var stylesXml = ReadZipEntry(bytes, "word/styles.xml");
        var documentXml = ReadZipEntry(bytes, "word/document.xml");

        Assert.Contains("""<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/>""", stylesXml);
        Assert.Contains("""<w:sz w:val="20"/><w:szCs w:val="20"/>""", stylesXml);
        Assert.Contains("""<w:color w:val="388E3C"/>""", stylesXml);
        Assert.Contains("""<w:color w:val="D32F2F"/>""", documentXml);
        Assert.Contains("""<w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="388E3C"/>""", documentXml);
        Assert.Contains("""<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="388E3C"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="C8E6C9"/>""", documentXml);
        Assert.Contains("""<w:trHeight w:val="480" w:hRule="atLeast"/>""", documentXml);
        Assert.Contains("""<w:trHeight w:val="400" w:hRule="atLeast"/>""", documentXml);
    }

    [Fact]
    public void Theme_CalledAfterPage_AppliesToHyperlinksAndTablesAddedAfterward()
    {
        IContainer? content = null;

        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                content = p.Content();
                content.Text("Before theme change.");
            });

            c.Theme(t => t
                .HyperlinkColor(Colors.Orange.L700)
                .TableHeaderBackground(Colors.Orange.L700));

            // Added to the page created above, but after Theme() - should pick up the new colors.
            content!.Text(t => t.Hyperlink("Link", "https://example.com"));
            content!.Table(table =>
            {
                table.HeaderRow(row => row.Cell().Text("Header"));
                table.Row(row => row.Cell().Text("Cell"));
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:color w:val="F57C00"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="F57C00"/>""", documentXml);
    }

    [Fact]
    public void PageLayout_SupportsFirstAndEvenHeadersFootersLandscapeAndPageNumberStart()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Landscape();
                p.PageNumberStart(5);
                p.FirstPageHeader().Text("First header");
                p.Header().Text("Odd header");
                p.EvenPageHeader().Text("Even header");
                p.FirstPageFooter().Text("First footer");
                p.Footer().Text("Odd footer");
                p.EvenPageFooter().Text("Even footer");
                p.Content().Text("Body.");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var documentRels = ReadZipEntry(bytes, "word/_rels/document.xml.rels");
        var contentTypes = ReadZipEntry(bytes, "[Content_Types].xml");
        var settingsXml = ReadZipEntry(bytes, "word/settings.xml");

        Assert.Contains("<w:headerReference w:type=\"default\"", documentXml);
        Assert.Contains("<w:headerReference w:type=\"first\"", documentXml);
        Assert.Contains("<w:headerReference w:type=\"even\"", documentXml);
        Assert.Contains("<w:footerReference w:type=\"default\"", documentXml);
        Assert.Contains("<w:footerReference w:type=\"first\"", documentXml);
        Assert.Contains("<w:footerReference w:type=\"even\"", documentXml);
        Assert.Contains("""<w:pgSz w:w="16837" w:h="11905"/>""", documentXml);
        Assert.Contains("""<w:pgNumType w:start="5"/>""", documentXml);
        Assert.Contains("""<w:titlePg/>""", documentXml);
        Assert.Contains("<w:evenAndOddHeaders/>", settingsXml);
        Assert.Equal(3, CountOccurrences(documentRels, "relationships/header"));
        Assert.Equal(3, CountOccurrences(documentRels, "relationships/footer"));
        Assert.Equal(3, CountOccurrences(contentTypes, "word/header"));
        Assert.Equal(3, CountOccurrences(contentTypes, "word/footer"));
    }

    [Fact]
    public void PageNumberFormat_EmitsWfmtAttribute()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.PageNumberStart(1);
                p.PageNumberFormat("lowerRoman");
                p.Content().H1("Roman Numeral Front Matter");
                p.Content().Text("This section uses i, ii, iii...");
            });
        }).PublishDocx();

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:pgNumType w:start="1" w:fmt="lowerRoman"/>""", documentXml);
    }

    [Fact]
    public void PageNumberFormat_WithoutStartValue()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.PageNumberFormat("decimal");
                p.Content().H1("Body with Explicit Decimal Format");
                p.Footer().Text(t =>
                {
                    t.Span("Page ");
                    t.CurrentPageNumber();
                });
            });
        }).PublishDocx();

        // Note: OpenXML SDK may not support pgNumType/@fmt validation yet;
        // verify XML generation instead.
        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:pgNumType w:fmt="decimal"/>""", documentXml);
    }

    [Fact]
    public void PageLayout_MultipleSectionsPreserveEachSectionsProperties()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.FirstPageHeader().Text("First section first header");
                p.EvenPageHeader().Text("First section even header");
                p.Header().Text("First section odd header");
                p.Content().Text("Portrait section.");
            });

            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Landscape();
                p.Header().Text("Landscape section header");
                p.Content().Text("Landscape section.");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("<w:headerReference w:type=\"first\"", documentXml);
        Assert.Contains("<w:headerReference w:type=\"even\"", documentXml);
        Assert.Contains("""<w:pgSz w:w="11905" w:h="16837"/>""", documentXml);
        Assert.Contains("""<w:pgSz w:w="16837" w:h="11905"/>""", documentXml);
    }

    [Fact]
    public void PageLayout_HeaderFooterWatermarkDoNotBleedIntoSectionWithoutOwnConfig()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Watermark("CONFIDENTIAL");
                p.Header().Text("Page 1 Header");
                p.Footer().Text("Page 1 Footer");
                p.Content().Text("Section 1 content.");
            });

            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text("Section 2 content.");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Equal(2, CountOccurrences(documentXml, "<w:headerReference w:type=\"default\""));
        Assert.Equal(2, CountOccurrences(documentXml, "<w:footerReference w:type=\"default\""));

        var header1 = ReadZipEntry(bytes, "word/header1.xml");
        var footer1 = ReadZipEntry(bytes, "word/footer1.xml");
        Assert.Contains("Page 1 Header", header1);
        Assert.Contains("CONFIDENTIAL", header1);
        Assert.Contains("Page 1 Footer", footer1);

        // Section 2 has no header/footer/watermark of its own, so TerraFluent.Docx.Reporting must emit
        // explicit empty overrides rather than letting section 2 inherit section 1's.
        var header2 = ReadZipEntry(bytes, "word/header2.xml");
        var footer2 = ReadZipEntry(bytes, "word/footer2.xml");
        Assert.DoesNotContain("Page 1 Header", header2);
        Assert.DoesNotContain("CONFIDENTIAL", header2);
        Assert.DoesNotContain("Page 1 Footer", footer2);
    }

    [Fact]
    public void PageLayout_EvenHeaderDoesNotBleedIntoSectionWithoutOwnEvenHeader()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Header().Text("Section 1 default header");
                p.EvenPageHeader().Text("Section 1 even header");
                p.Content().Text("Section 1 content.");
            });

            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Header().Text("Section 2 default header");
                p.Content().Text("Section 2 content.");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var settingsXml = ReadZipEntry(bytes, "word/settings.xml");

        Assert.Contains("<w:evenAndOddHeaders/>", settingsXml);
        Assert.Equal(2, CountOccurrences(documentXml, "<w:headerReference w:type=\"even\""));
        Assert.DoesNotContain("<w:titlePg/>", documentXml);

        var header2 = ReadZipEntry(bytes, "word/header2.xml");
        Assert.Contains("Section 1 even header", header2);

        // Section 2 doesn't define its own even-page header, so it must get an explicit
        // empty override instead of inheriting section 1's even header.
        var header4 = ReadZipEntry(bytes, "word/header4.xml");
        Assert.DoesNotContain("Section 1 even header", header4);
    }

    [Fact]
    public void NamedStyles_ApplyParagraphAndTableStyles()
    {
        var bytes = Document.Create(c =>
        {
            c.ParagraphStyle("Executive Callout", style => style
                .Bold()
                .FontColor(Colors.Blue.L800)
                .FontSize(12)
                .LeftIndent(18)
                .SpacingBefore(6)
                .SpacingAfter(6));

            c.TableStyle("Financial Table", table => table
                .Border(0.75f, Colors.Blue.L700)
                .CellPadding(5, 7)
                .HeaderBackground(Colors.Blue.L700)
                .AlternateRowBackground(Colors.Grey.L100)
                .HeaderRowMinHeight(24)
                .RowMinHeight(20));

            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text("Styled paragraph.", t => t.Style("Executive Callout"));
                p.Content().Table(table =>
                {
                    table.Style("Financial Table");
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn();
                        cols.RelativeColumn();
                    });
                    table.HeaderRow(row =>
                    {
                        row.Cell().Text("Metric", t => t.Bold().FontColor(Colors.White.Default));
                        row.Cell().Text("Value", t => t.Bold().FontColor(Colors.White.Default));
                    });
                    table.Row(row => { row.Cell().Text("Revenue"); row.Cell().Text("$10M"); });
                    table.Row(row => { row.Cell().Text("Margin"); row.Cell().Text("20%"); });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var stylesXml = ReadZipEntry(bytes, "word/styles.xml");

        Assert.Contains("""<w:pStyle w:val="ExecutiveCallout"/>""", documentXml);
        Assert.Contains("""<w:tblStyle w:val="FinancialTable"/>""", documentXml);
        Assert.Contains("""<w:style w:type="paragraph" w:customStyle="1" w:styleId="ExecutiveCallout">""", stylesXml);
        Assert.Contains("""<w:name w:val="Executive Callout"/>""", stylesXml);
        Assert.Contains("""<w:style w:type="table" w:customStyle="1" w:styleId="FinancialTable">""", stylesXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="1976D2"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>""", documentXml);
    }

    [Fact]
    public void TocBookmarksAndCrossReferences_EmitFieldsAndAnchors()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().TableOfContents("Contents", minLevel: 1, maxLevel: 2);
                p.Content().H1("Executive Summary");
                p.Content().Bookmark("revenue-section", "Revenue Performance", t => t.Style("Heading2"));
                p.Content().Text(t =>
                {
                    t.Span("See ");
                    t.CrossReference("revenue-section", "Revenue Performance");
                    t.Span(" for details.");
                });
                p.Content().Bookmark("empty-anchor");
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");

        Assert.Contains(" TOC \\o &quot;1-2&quot; \\h \\z \\u ", documentXml);
        Assert.DoesNotContain("w:dirty=\"true\"", documentXml);
        Assert.Contains("<w:bookmarkStart w:id=\"", documentXml);
        Assert.Contains("w:name=\"revenue_section\"", documentXml);
        Assert.Contains("w:name=\"empty_anchor\"", documentXml);
        Assert.Contains(" REF revenue_section \\h ", documentXml);
        Assert.Contains("Revenue Performance", documentXml);
    }

    [Fact]
    public void AdvancedTables_EmitWidthAlignmentNoSplitBordersMergesAndTextDirection()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Table(table =>
                {
                    table.WidthPercent(80).AlignCenter().Border(0.5f, Colors.Grey.L400);
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn();
                        cols.RelativeColumn();
                        cols.RelativeColumn();
                    });
                    table.HeaderRow(row =>
                    {
                        row.KeepTogether();
                        row.Cell().Text("Group");
                        row.Cell().Text("Item");
                        row.Cell().Text("Value");
                    });
                    table.Row(row =>
                    {
                        row.KeepTogether();
                        row.Cell().VerticalMergeStart().Border(1, Colors.Blue.L700).Text("Revenue");
                        row.Cell().Text("Services");
                        row.Cell().Text("$10M", t => t.AlignRight());
                    });
                    table.Row(row =>
                    {
                        row.Cell().VerticalMergeContinue();
                        row.Cell().Text("Licensing");
                        row.Cell().Text("$2M", t => t.AlignRight());
                    });
                    table.Row(row =>
                    {
                        row.Cell(2).Text("Vertical label");
                        row.Cell().TextDirectionTopToBottom().Text("Approved");
                    });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:tblW w:w="4000" w:type="pct"/>""", documentXml);
        Assert.Contains("""<w:jc w:val="center"/>""", documentXml);
        Assert.Contains("<w:cantSplit/>", documentXml);
        Assert.Contains("""<w:vMerge w:val="restart"/>""", documentXml);
        Assert.Contains("<w:vMerge/>", documentXml);
        Assert.Contains("""<w:tcBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="1976D2"/>""", documentXml);
        Assert.Contains("""<w:textDirection w:val="tbRl"/>""", documentXml);
    }

    [Fact]
    public void Table_WithoutColumnsDefinition_InfersColumnsFromWidestRow()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Table(table =>
                {
                    table.HeaderRow(row =>
                    {
                        row.Cell().Text("Group");
                        row.Cell().Text("Item");
                        row.Cell().Text("Value");
                    });
                    table.Row(row =>
                    {
                        row.Cell().Text("Revenue");
                        row.Cell(2).Text("Services and Licensing");
                    });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Equal(3, CountOccurrences(documentXml, "<w:gridCol"));
        Assert.Contains("Group", documentXml);
        Assert.Contains("Item", documentXml);
        Assert.Contains("Value", documentXml);
        Assert.Contains("Revenue", documentXml);
        Assert.Contains("""<w:gridSpan w:val="2"/>""", documentXml);
        Assert.Contains("Services and Licensing", documentXml);
    }

    [Fact]
    public void Table_RowExceedsDefinedColumns_Throws()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            Document.Create(c =>
            {
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Content().Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn();
                            cols.RelativeColumn();
                        });
                        table.Row(row =>
                        {
                            row.Cell().Text("A");
                            row.Cell().Text("B");
                            row.Cell().Text("C");
                        });
                    });
                });
            }).PublishDocx());

        Assert.Contains("row 1", ex.Message);
        Assert.Contains("3 column(s)", ex.Message);
        Assert.Contains("2 column(s)", ex.Message);
    }

    [Fact]
    public void Table_MixedRelativeAndConstantColumns_WithoutWidth_UsesDefaultPageWidthReference()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.ConstantColumn(100);
                        cols.RelativeColumn();
                        cols.RelativeColumn();
                    });
                    table.HeaderRow(row =>
                    {
                        row.Cell().Text("Fixed");
                        row.Cell().Text("Flex A");
                        row.Cell().Text("Flex B");
                    });
                    table.Row(row =>
                    {
                        row.Cell().Text("100pt");
                        row.Cell().Text("share");
                        row.Cell().Text("share");
                    });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");

        // gridCol widths are all dxa (twentieths of a point): the 100pt constant column is
        // 2000, and each relative column gets half of the remaining 368pt (out of an
        // assumed 468pt page width) = 184pt = 3680.
        Assert.Equal(1, CountOccurrences(documentXml, """<w:gridCol w:w="2000"/>"""));
        Assert.Equal(2, CountOccurrences(documentXml, """<w:gridCol w:w="3680"/>"""));

        // tcW widths are all pct (out of 5000) for every cell, including the constant column,
        // instead of mixing dxa and pct units in the same row.
        Assert.Equal(2, CountOccurrences(documentXml, """<w:tcW w:w="1068" w:type="pct"/>"""));
        Assert.Equal(4, CountOccurrences(documentXml, """<w:tcW w:w="1965" w:type="pct"/>"""));
    }

    [Fact]
    public void Table_MixedRelativeAndConstantColumns_WithExplicitWidth_UsesItAsReference()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Table(table =>
                {
                    table.Width(360);
                    table.ColumnsDefinition(cols =>
                    {
                        cols.ConstantColumn(60);
                        cols.RelativeColumn();
                        cols.RelativeColumn();
                    });
                    table.Row(row =>
                    {
                        row.Cell().Text("60pt");
                        row.Cell().Text("share");
                        row.Cell().Text("share");
                    });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");

        // Table width is the explicit 360pt (7200 dxa); gridCol widths split it the same way
        // (60pt constant = 1200 dxa, remaining 300pt split evenly = 150pt = 3000 dxa each).
        Assert.Contains("""<w:tblW w:w="7200" w:type="dxa"/>""", documentXml);
        Assert.Equal(1, CountOccurrences(documentXml, """<w:gridCol w:w="1200"/>"""));
        Assert.Equal(2, CountOccurrences(documentXml, """<w:gridCol w:w="3000"/>"""));

        // tcW widths are pct of the explicit table width, consistently for constant and
        // relative columns alike.
        Assert.Equal(1, CountOccurrences(documentXml, """<w:tcW w:w="833" w:type="pct"/>"""));
        Assert.Equal(2, CountOccurrences(documentXml, """<w:tcW w:w="2083" w:type="pct"/>"""));
    }

    [Fact]
    public void HexColor_AcceptsLeadingHashAndNormalizesToUppercase()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text(t => t.Span("Hello", x => x.FontColor("#1976d2")));
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:color w:val="1976D2"/>""", documentXml);
    }

    [Fact]
    public void HexColor_ThreeDigitShorthand_ExpandsToSixDigits()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text(t => t.Span("Hello", x => x.FontColor("D32")));
                p.Content().Table(table =>
                {
                    table.HeaderBackground("0F0");
                    table.HeaderRow(row => row.Cell().Text("Header"));
                    table.Row(row => row.Cell().Text("Body"));
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:color w:val="DD3322"/>""", documentXml);
        Assert.Contains("""<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>""", documentXml);
    }

    [Fact]
    public void HexColor_InvalidValue_ThrowsArgumentException()
    {
        var wrongLength = Assert.Throws<ArgumentException>(() =>
            Document.Create(c => c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text(t => t.Span("Hello", x => x.FontColor("12345")));
            })).PublishDocx());
        Assert.Contains("hexColor", wrongLength.Message);
        Assert.Contains("12345", wrongLength.Message);

        var nonHexDigits = Assert.Throws<ArgumentException>(() =>
            Document.Create(c => c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text("Hello", t => t.Shading("ZZZZZZ"));
            })).PublishDocx());
        Assert.Contains("hexColor", nonHexDigits.Message);
        Assert.Contains("ZZZZZZ", nonHexDigits.Message);

        var emptyValue = Assert.Throws<ArgumentException>(() =>
            Document.Create(c => c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Background("");
            })).PublishDocx());
        Assert.Contains("hexColor", emptyValue.Message);
    }

    [Fact]
    public void TabStops_EmitParagraphTabsAndTabRuns()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text(t =>
                {
                    t.TabStop(240, TabStopAlignment.Right);
                    t.TabStop(320, TabStopAlignment.Decimal);
                    t.Span("Subtotal");
                    t.Tab();
                    t.Span("$1,234.50");
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        Assert.Contains("""<w:tabs><w:tab w:val="right" w:pos="4800"/><w:tab w:val="decimal" w:pos="6400"/></w:tabs>""", documentXml);
        Assert.Contains("<w:tab/>", documentXml);
    }

    [Fact]
    public void FootnotesAndEndnotes_WriteNotePartsAndReferences()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text(t =>
                {
                    t.Span("Statement with a footnote");
                    t.Footnote("Footnote detail.");
                    t.Span(" and an endnote");
                    t.Endnote("Endnote detail.");
                    t.Span(".");
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var relsXml = ReadZipEntry(bytes, "word/_rels/document.xml.rels");
        var contentTypes = ReadZipEntry(bytes, "[Content_Types].xml");
        var footnotesXml = ReadZipEntry(bytes, "word/footnotes.xml");
        var endnotesXml = ReadZipEntry(bytes, "word/endnotes.xml");

        Assert.Contains("""<w:footnoteReference w:id="1"/>""", documentXml);
        Assert.Contains("""<w:endnoteReference w:id="1"/>""", documentXml);
        Assert.Contains("relationships/footnotes", relsXml);
        Assert.Contains("relationships/endnotes", relsXml);
        Assert.Contains("footnotes+xml", contentTypes);
        Assert.Contains("endnotes+xml", contentTypes);
        Assert.Contains("Footnote detail.", footnotesXml);
        Assert.Contains("Endnote detail.", endnotesXml);
    }

    [Fact]
    public void Footnote_WithHyperlink_RendersRelationshipAndRuns()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Text(t =>
                {
                    t.Span("Statement with a linked footnote");
                    t.Footnote("See ", ft => ft.Hyperlink("our website", "https://example.com/footnote-link"));
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var footnotesXml = ReadZipEntry(bytes, "word/footnotes.xml");
        var footnoteRels = ReadZipEntry(bytes, "word/_rels/footnotes.xml.rels");

        Assert.Contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", footnotesXml);
        Assert.Contains("""<w:hyperlink r:id="rId1" w:history="1">""", footnotesXml);
        Assert.Contains("our website", footnotesXml);
        Assert.Contains("relationships/hyperlink", footnoteRels);
        Assert.Contains("https://example.com/footnote-link", footnoteRels);
    }

    [Fact]
    public void Endnote_WithCrossReferenceAndPageFields_RendersFieldCharSequences()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Margin(Unit.Centimetre(2.54f));
                p.Content().Bookmark("appendix", "Appendix", t => t.Style("Heading2"));
                p.Content().Text(t =>
                {
                    t.Span("Detail in appendix");
                    t.Endnote("See ", en =>
                    {
                        en.CrossReference("appendix", "Appendix");
                        en.Span(" on page ");
                        en.CurrentPageNumber();
                        en.Span(" of ");
                        en.TotalPages();
                    });
                });
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var endnotesXml = ReadZipEntry(bytes, "word/endnotes.xml");

        Assert.Contains(" REF appendix \\h ", endnotesXml);
        Assert.Contains("Appendix", endnotesXml);
        Assert.Contains(" PAGE ", endnotesXml);
        Assert.Contains(" NUMPAGES ", endnotesXml);
        Assert.Contains("""<w:fldChar w:fldCharType="begin"/>""", endnotesXml);
        Assert.Contains("""<w:fldChar w:fldCharType="separate"/>""", endnotesXml);
        Assert.Contains("""<w:fldChar w:fldCharType="end"/>""", endnotesXml);
    }

    [Fact]
    public void NestedFootnoteOrEndnote_Throws()
    {
        var footnoteInFootnote = Assert.Throws<InvalidOperationException>(() =>
            Document.Create(c =>
            {
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Content().Text(t =>
                    {
                        t.Span("Statement");
                        t.Footnote("Outer note", ft => ft.Footnote("Inner note"));
                    });
                });
            }));
        Assert.Contains("nested footnote or endnote", footnoteInFootnote.Message);

        var footnoteInEndnote = Assert.Throws<InvalidOperationException>(() =>
            Document.Create(c =>
            {
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Content().Text(t =>
                    {
                        t.Span("Statement");
                        t.Endnote("Outer note", en => en.Footnote("Inner note"));
                    });
                });
            }));
        Assert.Contains("nested footnote or endnote", footnoteInEndnote.Message);
    }

    [Fact]
    public void ChartsBackgroundAndWatermark_EmitDocumentContent()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Background(Colors.Grey.L100);
                p.Watermark("DRAFT");
                p.Content().Chart(chart => chart
                    .Title("Revenue Trend")
                    .Series(s => s
                        .Color(Colors.Green.L700)
                        .Bar("Q1", 4.2)
                        .Bar("Q2", 4.8)
                        .Bar("Q3", 5.1)));
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var documentXml = ReadZipEntry(bytes, "word/document.xml");
        var relsXml = ReadZipEntry(bytes, "word/_rels/document.xml.rels");
        var contentTypes = ReadZipEntry(bytes, "[Content_Types].xml");
        var settingsXml = ReadZipEntry(bytes, "word/settings.xml");
        var headerXml = ReadZipEntry(bytes, "word/header1.xml");
        var chartXml = ReadZipEntry(bytes, "word/charts/chart1.xml");

        Assert.Contains("""<w:background w:color="F5F5F5">""", documentXml);
        Assert.Contains("""<v:fill color2="#F5F5F5" type="solid"/>""", documentXml);
        Assert.Contains("<w:displayBackgroundShape/>", settingsXml);
        Assert.Contains("DRAFT", headerXml);
        Assert.Contains("rotation:315", headerXml);
        Assert.Contains("z-index:-251654144", headerXml);
        Assert.Contains("""<v:fill opacity=".55"/>""", headerXml);
        Assert.Contains("<v:textpath", headerXml);
        Assert.Contains("drawingml/2006/chart", documentXml);
        Assert.Contains("relationships/chart\" Target=\"charts/chart1.xml\"", relsXml);
        Assert.Contains("drawingml.chart+xml", contentTypes);
        Assert.Contains("Revenue Trend", chartXml);
        Assert.Contains("Q1", chartXml);
        Assert.Contains("<c:barChart>", chartXml);
    }

    [Fact]
    public void MultiSeriesBarChart_GeneratesMultipleSerElements()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Chart(chart => chart
                    .Title("Revenue Comparison")
                    .Series("2024", s => s.Bar("Q1", 100).Bar("Q2", 120).Bar("Q3", 140))
                    .Series("2025", s => s.Bar("Q1", 110).Bar("Q2", 130).Bar("Q3", 150)));
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var chartXml = ReadZipEntry(bytes, "word/charts/chart1.xml");

        // Should have 2 series elements with idx 0 and 1
        Assert.Contains("""<c:idx val="0"/>""", chartXml);
        Assert.Contains("""<c:idx val="1"/>""", chartXml);
        Assert.Contains("""<c:order val="0"/>""", chartXml);
        Assert.Contains("""<c:order val="1"/>""", chartXml);

        // Series names in legend
        Assert.Contains("""<c:tx><c:v>2024</c:v></c:tx>""", chartXml);
        Assert.Contains("""<c:tx><c:v>2025</c:v></c:tx>""", chartXml);

        // Should still be a bar chart
        Assert.Contains("<c:barChart>", chartXml);
    }

    [Fact]
    public void MultiSeriesChart_EachSeriesHasOwnColor()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Chart(chart => chart
                    .Title("Revenue vs Cost")
                    .Series("Revenue", s => s.Color("00AA00").Bar("Q1", 100).Bar("Q2", 120))
                    .Series("Cost", s => s.Color("FF0000").Bar("Q1", 60).Bar("Q2", 70)));
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var chartXml = ReadZipEntry(bytes, "word/charts/chart1.xml");

        // Both colors should be present in the chart
        Assert.Contains("""00AA00""", chartXml);
        Assert.Contains("""FF0000""", chartXml);
    }

    [Fact]
    public void SingleSeriesChart_NewAPIStillWorks()
    {
        var bytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Chart(chart => chart
                    .Title("Sales")
                    .Series(s => s.Line("Jan", 10).Line("Feb", 15).Line("Mar", 12)));
            });
        }).PublishDocx();

        Assert.Empty(Validate(bytes));

        var chartXml = ReadZipEntry(bytes, "word/charts/chart1.xml");

        // Should have exactly one series
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(chartXml, """<c:idx val="0"/>"""));
        Assert.Contains("Jan", chartXml);
        Assert.Contains("<c:lineChart>", chartXml);
    }

    [Fact]
    public void Chart_MixingSeriesTypes_Throws()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            Document.Create(c =>
            {
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Content().Chart(chart => chart
                        .Title("Mixed")
                        .Series(s => s.Bar("Q1", 1).Bar("Q2", 2))
                        .Series(s => s.Line("Q1", 3).Line("Q2", 4)));
                });
            }));

        Assert.Contains("bar", ex.Message);
        Assert.Contains("line", ex.Message);
    }

    [Fact]
    public void Chart_MultipleSeriesWithDataOnPieChart_Throws()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            Document.Create(c =>
            {
                c.Page(p =>
                {
                    p.Size(PageSize.A4);
                    p.Content().Chart(chart => chart
                        .Title("Share")
                        .Series(s => s.Pie("A", 1).Pie("B", 2))
                        .Series(s => s.Pie("C", 3)));
                });
            }));

        Assert.Contains("single data series", ex.Message);
    }

    [Fact]
    public void DocxTemplate_ReplacesPlaceholders()
    {
        var templateBytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text("Hello {{Name}}.");
            });
        }).PublishDocx();

        var templatePath = Path.Combine(Path.GetTempPath(), $"template-{Guid.NewGuid():N}.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"template-output-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(templatePath, templateBytes);

        try
        {
            DocxTemplate.Open(templatePath)
                .Replace("Name", "Ada Lovelace")
                .SaveAs(outputPath);

            var bytes = File.ReadAllBytes(outputPath);
            Assert.Empty(Validate(bytes));
            var documentXml = ReadZipEntry(bytes, "word/document.xml");
            Assert.Contains("Hello Ada Lovelace.", documentXml);
            Assert.DoesNotContain("{{Name}}", documentXml);
        }
        finally
        {
            File.Delete(templatePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void DocxTemplate_ReplacesSplitRunPlaceholdersWithoutBareKeyReplacement()
    {
        var templateBytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text(t =>
                {
                    t.Span("Hello {{Na");
                    t.Span("me}} and Name.");
                });
            });
        }).PublishDocx();

        var templatePath = Path.Combine(Path.GetTempPath(), $"template-{Guid.NewGuid():N}.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"template-output-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(templatePath, templateBytes);

        try
        {
            DocxTemplate.Open(templatePath)
                .Replace("Name", "Ada Lovelace")
                .SaveAs(outputPath);

            var bytes = File.ReadAllBytes(outputPath);
            Assert.Empty(Validate(bytes));
            var documentText = PlainText(ReadZipEntry(bytes, "word/document.xml"));
            Assert.Contains("Hello Ada Lovelace and Name.", documentText);
            Assert.DoesNotContain("{{Name}}", documentText);
        }
        finally
        {
            File.Delete(templatePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void DocxTemplate_ReplacesTaggedContentControl()
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var baseBytes = Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Text("Old Customer");
            });
        }).PublishDocx();

        var templateBytes = RewriteZipEntry(baseBytes, "word/document.xml", xml =>
        {
            var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            var paragraph = doc.Descendants(w + "p").First();
            paragraph.ReplaceNodes(
                new XElement(w + "sdt",
                    new XElement(w + "sdtPr",
                        new XElement(w + "tag", new XAttribute(w + "val", "CustomerName"))),
                    new XElement(w + "sdtContent",
                        new XElement(w + "r",
                            new XElement(w + "t", "Old Customer")))));
            return doc.ToString(SaveOptions.DisableFormatting);
        });

        var templatePath = Path.Combine(Path.GetTempPath(), $"template-{Guid.NewGuid():N}.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"template-output-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(templatePath, templateBytes);

        try
        {
            DocxTemplate.Open(templatePath)
                .ReplaceContentControl("CustomerName", "Ada & Co")
                .SaveAs(outputPath);

            var bytes = File.ReadAllBytes(outputPath);
            Assert.Empty(Validate(bytes));
            var documentText = PlainText(ReadZipEntry(bytes, "word/document.xml"));
            Assert.Contains("Ada & Co", documentText);
            Assert.DoesNotContain("Old Customer", documentText);
        }
        finally
        {
            File.Delete(templatePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void PublicApis_InvalidInputsThrowClearExceptions()
    {
        Assert.Throws<ArgumentNullException>(() => Document.Create(null!));

        Assert.Throws<ArgumentOutOfRangeException>(() => Document.Create(c =>
            c.Page(p => p.Size(-1, PageSize.A4.Height))));

        Assert.Throws<ArgumentOutOfRangeException>(() => Document.Create(c =>
            c.Page(p => p.Margin(-1))));

        Assert.Throws<ArgumentOutOfRangeException>(() => Document.Create(c =>
            c.Page(p => p.Columns(0))));

        Assert.Throws<ArgumentException>(() => Document.Create(c =>
            c.Page(p => p.Content().Image(Array.Empty<byte>(), "empty.png"))));

        var missingImage = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.png");
        Assert.Throws<FileNotFoundException>(() => Document.Create(c =>
        {
            c.Page(p =>
            {
                p.Size(PageSize.A4);
                p.Content().Image(missingImage);
            });
        }).PublishDocx());
    }

    private static byte[] CreateMinimalPng(int width, int height)
    {
        // Minimal valid PNG: signature + IHDR + IDAT (1px red) + IEND
        using var ms = new MemoryStream();
        // PNG signature
        ms.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        // IHDR chunk (13 bytes data)
        WriteChunk(ms, "IHDR", [
            (byte)(width >> 24), (byte)(width >> 16), (byte)(width >> 8), (byte)width,
            (byte)(height >> 24), (byte)(height >> 16), (byte)(height >> 8), (byte)height,
            8, 2, 0, 0, 0  // 8-bit depth, RGB color, compression=0, filter=0, interlace=0
        ]);
        // IDAT chunk: zlib-compressed scanline (filter byte 0 + R G B)
        var scanline = new byte[] { 0, 255, 0, 0 }; // filter=None, red pixel
        var compressed = ZlibCompress(scanline);
        WriteChunk(ms, "IDAT", compressed);
        // IEND chunk
        WriteChunk(ms, "IEND", []);
        return ms.ToArray();
    }

    private static void WriteChunk(MemoryStream ms, string type, byte[] data)
    {
        var len = BitConverter.GetBytes(data.Length);
        if (BitConverter.IsLittleEndian) Array.Reverse(len);
        ms.Write(len);
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        ms.Write(typeBytes);
        ms.Write(data);
        var crcData = new byte[4 + data.Length];
        typeBytes.CopyTo(crcData, 0);
        data.CopyTo(crcData, 4);
        var crc = Crc32(crcData);
        var crcBytes = BitConverter.GetBytes(crc);
        if (BitConverter.IsLittleEndian) Array.Reverse(crcBytes);
        ms.Write(crcBytes);
    }

    private static byte[] ZlibCompress(byte[] data)
    {
        using var output = new MemoryStream();
        output.WriteByte(0x78); output.WriteByte(0x9C); // zlib header
        using (var deflate = new System.IO.Compression.DeflateStream(output, System.IO.Compression.CompressionMode.Compress, true))
            deflate.Write(data, 0, data.Length);
        var adler = Adler32(data);
        var adlerBytes = BitConverter.GetBytes(adler);
        if (BitConverter.IsLittleEndian) Array.Reverse(adlerBytes);
        output.Write(adlerBytes);
        return output.ToArray();
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }
        return ~crc;
    }

    private static uint Adler32(byte[] data)
    {
        uint s1 = 1, s2 = 0;
        foreach (var b in data) { s1 = (s1 + b) % 65521; s2 = (s2 + s1) % 65521; }
        return (s2 << 16) | s1;
    }

    private static List<string> Validate(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        return validator.Validate(doc)
            .Select(e => $"[{e.ErrorType}] {e.Description} (path: {e.Path?.XPath})")
            .ToList();
    }

    private static string ReadZipEntry(byte[] bytes, string path)
    {
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = archive.GetEntry(path) ?? throw new FileNotFoundException(path);
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static byte[] RewriteZipEntry(byte[] bytes, string path, Func<string, string> rewrite)
    {
        using var sourceStream = new MemoryStream(bytes);
        using var source = new ZipArchive(sourceStream, ZipArchiveMode.Read);
        using var outputStream = new MemoryStream();
        using (var output = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in source.Entries)
            {
                var newEntry = output.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var entryStream = entry.Open();
                using var newEntryStream = newEntry.Open();

                if (entry.FullName.Equals(path, StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    using var writer = new StreamWriter(newEntryStream, new UTF8Encoding(false));
                    writer.Write(rewrite(reader.ReadToEnd()));
                }
                else
                {
                    entryStream.CopyTo(newEntryStream);
                }
            }
        }

        return outputStream.ToArray();
    }

    private static string PlainText(string xml)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var document = XDocument.Parse(xml);
        return string.Concat(document.Descendants(w + "t").Select(t => t.Value));
    }

    private static int CountOccurrences(string text, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }
        return count;
    }
}
