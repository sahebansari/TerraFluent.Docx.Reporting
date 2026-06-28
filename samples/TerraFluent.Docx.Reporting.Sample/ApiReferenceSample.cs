using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;

internal static class ApiReferenceSample
{
    public static string Generate(string outputDirectory, string imagePath)
    {
        var documentPath = Path.Combine(outputDirectory, "sample-api-reference.docx");
        var imageBytes = File.ReadAllBytes(imagePath);

        Document.Create(container =>
        {
            container
                .Theme(theme => theme
                    .DefaultFont("Aptos", 10.5f)
                    .DefaultTextColor(Colors.Grey.L900)
                    .HeadingColor(Colors.Blue.L800)
                    .AccentColor(Colors.Blue.L700)
                    .HyperlinkColor(Colors.Red.L700)
                    .TableBorder(0.75f, Colors.Grey.L300)
                    .TableCellPadding(4, 6)
                    .TableHeaderBackground(Colors.Blue.L700)
                    .TableAlternateRowBackground(Colors.Grey.L100)
                    .TableRowMinHeight(20)
                    .TableHeaderRowMinHeight(22))
                .ParagraphStyle("Callout", style => style
                    .Italic()
                    .FontColor(Colors.Blue.L800)
                    .Shading(Colors.Blue.L100)
                    .LeftIndent(12)
                    .RightIndent(12)
                    .SpacingBefore(4)
                    .SpacingAfter(4)
                    .KeepLinesTogether())
                .TableStyle("Reference Table", table => table
                    .Border(0.75f, Colors.Grey.L300)
                    .CellPadding(4, 6)
                    .HeaderBackground(Colors.Blue.L700)
                    .AlternateRowBackground(Colors.Grey.L100)
                    .HeaderRowMinHeight(22)
                    .RowMinHeight(20))
                .MetadataTitle("TerraFluent.Docx.Reporting API Reference")
                .MetadataAuthor("TerraFluent.Docx.Reporting")
                .MetadataSubject("A systematic tour of every TerraFluent.Docx.Reporting descriptor API")
                .MetadataKeywords("TerraFluent.Docx.Reporting, reference, API, docx, sample")
                .MetadataCreator("TerraFluent.Docx.Reporting Sample App");

            // ----------------------------------------------------------------
            // Page 1 - Cover, Table of Contents, Document Setup, Page Setup
            // ----------------------------------------------------------------
            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.PageNumberStart(1);
                page.PageNumberFormat("lowerRoman");
                page.Background(Colors.Grey.L100);
                page.Watermark("SAMPLE", Colors.Grey.L300, 90);
                page.DefaultTextStyle(t => t.FontFamily("Aptos").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(4));

                page.FirstPageHeader().Text("TerraFluent.Docx.Reporting API Reference - Cover", t => t.Bold().FontColor(Colors.Blue.L800).AlignCenter());

                page.Header().Row(row =>
                {
                    row.RelativeItem().Text("TerraFluent.Docx.Reporting API Reference", t => t.Bold().FontColor(Colors.Blue.L800).FontSize(9));
                    row.AutoItem().Text(t => t.CurrentPageNumber(s => s.FontSize(9).FontColor(Colors.Grey.L600)).AlignRight());
                }).Line();

                page.EvenPageHeader().Row(row =>
                {
                    row.RelativeItem().Text(t => t.CurrentPageNumber().FontSize(9).FontColor(Colors.Grey.L600));
                    row.AutoItem().Text("TerraFluent.Docx.Reporting API Reference (even page)", t => t.Bold().FontColor(Colors.Grey.L700).FontSize(9).AlignRight());
                }).Line();

                page.FirstPageFooter().Text("Cover page footer (FirstPageFooter)", t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignCenter());
                page.Footer().Line().Text(t =>
                {
                    t.Span("Page ").FontSize(9).FontColor(Colors.Grey.L600);
                    t.CurrentPageNumber(s => s.FontSize(9).FontColor(Colors.Grey.L600));
                    t.Span(" of ").FontSize(9).FontColor(Colors.Grey.L600);
                    t.TotalPages(s => s.FontSize(9).FontColor(Colors.Grey.L600));
                    t.AlignCenter();
                });
                page.EvenPageFooter().Text("Even page footer (EvenPageFooter)", t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignCenter());

                page.Content().Column(col =>
                {
                    col.Spacing(8);

                    col.Item().H1("TerraFluent.Docx.Reporting API Reference");
                    col.Item().Text("This document is a systematic tour of every public TerraFluent.Docx.Reporting descriptor API: " +
                                     "document setup and theming, page setup, text and typography, lists, images, " +
                                     "tables, charts, and layout primitives. Each section pairs the API call with " +
                                     "its rendered output.", t => t.Justify().SpacingAfter(8));

                    col.Item().TableOfContents("Contents", minLevel: 1, maxLevel: 2);

                    col.Item().PageBreak();

                    col.Item().Bookmark("sec-document-setup", "1. Document Setup & Theme", t => t.Style("Heading1"));
                    col.Item().Text("Document.Create(configure) builds a document. The configure callback receives an " +
                                     "IDocumentContainer, which sets document-wide concerns:");
                    col.Item().BulletList(list =>
                    {
                        list.Item(t =>
                        {
                            t.Span("Theme(configure)").Bold();
                            t.Span(" - default font/size/colors, table border/padding/banding defaults. " +
                                   "Applied across all pages and table styles created afterward.");
                        });
                        list.Item(t =>
                        {
                            t.Span("ParagraphStyle(name, configure)").Bold();
                            t.Span(" - a named, reusable paragraph style. This document defines a \"Callout\" " +
                                   "style (italic, shaded, indented) used throughout.");
                        });
                        list.Item(t =>
                        {
                            t.Span("TableStyle(name, configure)").Bold();
                            t.Span(" - a named, reusable table style. This document defines a \"Reference Table\" " +
                                   "style applied with .Style(\"Reference Table\") in section 6.");
                        });
                        list.Item(t =>
                        {
                            t.Span("MetadataTitle / MetadataAuthor / MetadataSubject / MetadataKeywords / MetadataCreator").Bold();
                            t.Span(" - set the .docx core properties (visible in Word's File > Info panel).");
                        });
                    });

                    col.Item().Text(t =>
                    {
                        t.Span("This paragraph uses the \"Callout\" style defined above via ");
                        t.Span(".Style(\"Callout\")").Bold();
                        t.Span(". Named styles keep formatting consistent without repeating the same " +
                               "Shading/Border/Indent calls on every paragraph.");
                        t.Style("Callout");
                    });

                    col.Item().Bookmark("sec-page-setup", "2. Page Setup", t => t.Style("Heading1"));
                    col.Item().Text("Each container.Page(configure) call starts a new section with its own size, " +
                                     "margins, headers/footers, columns, and page-numbering rules.");
                    col.Item().BulletList(list =>
                    {
                        list.Item(t =>
                        {
                            t.Span("Size(PageSize.A4 / .Letter / .Legal / ...) or Size(width, height)").Bold();
                            t.Span(" plus .Landscape() / .Portrait() - this page uses A4 portrait.");
                        });
                        list.Item(t =>
                        {
                            t.Span("Margin(all) / Margin(vertical, horizontal) / Margin(top, right, bottom, left)").Bold();
                            t.Span(" and MarginTop/Right/Bottom/Left for individual edges.");
                        });
                        list.Item(t =>
                        {
                            t.Span("PageNumberStart(n) and PageNumberFormat(\"decimal\"|\"lowerRoman\"|\"upperRoman\"|\"lowerLetter\"|\"upperLetter\"|\"none\")").Bold();
                            t.Span(" - this cover page is numbered with lowercase roman numerals starting at i. " +
                                   "Section 3 onward restarts at 1 in Arabic numerals.");
                        });
                        list.Item(t =>
                        {
                            t.Span("Background(hexColor)").Bold();
                            t.Span(" - sets Word's document-wide page color (this page's light grey background). " +
                                   "Only the first non-empty value across all pages is used.");
                        });
                        list.Item(t =>
                        {
                            t.Span("Watermark(text, hexColor, fontSize)").Bold();
                            t.Span(" - the \"SAMPLE\" diagonal watermark behind this page's content.");
                        });
                        list.Item(t =>
                        {
                            t.Span("Columns(count, spacingPoints, separatorLine) and SingleColumn()").Bold();
                            t.Span(" - section 4 demonstrates a two-column newsletter layout.");
                        });
                        list.Item(t =>
                        {
                            t.Span("Header() / Footer() (aliases: OddPageHeader() / OddPageFooter()), FirstPageHeader() / FirstPageFooter(), and EvenPageHeader() / EvenPageFooter()").Bold();
                            t.Span(" - this cover page shows all three header and footer variants at once.");
                        });
                        list.Item(t =>
                        {
                            t.Span("DefaultTextStyle(configure)").Bold();
                            t.Span(" - sets the baseline font/size/color/spacing applied to text in this section.");
                        });
                    });
                });
            });

            // ----------------------------------------------------------------
            // Page 2 - Text & Typography
            // ----------------------------------------------------------------
            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.PageNumberStart(1);
                page.PageNumberFormat("decimal");
                page.DefaultTextStyle(t => t.FontFamily("Aptos").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(4));
                AddStandardHeaderFooter(page, "3. Text & Typography");

                page.Content().Column(col =>
                {
                    col.Spacing(6);

                    col.Item().Bookmark("sec-typography", "3. Text & Typography", t => t.Style("Heading1"));
                    col.Item().Text("ITextDescriptor configures paragraph- and run-level formatting. Use .Span(text, configure) " +
                                     "to format individual runs within a paragraph.");

                    col.Item().H2("3.1 Character Formatting");
                    col.Item().Text(t =>
                    {
                        t.Span("Bold").Bold();
                        t.Span(", ");
                        t.Span("italic").Italic();
                        t.Span(", ");
                        t.Span("underline").Underline();
                        t.Span(", ");
                        t.Span("strikethrough").Strikethrough();
                        t.Span(", and combined ").Bold();
                        t.Span("bold italic underline").Bold().Italic().Underline();
                        t.Span(".");
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("Superscript and subscript: E=mc");
                        t.Span("2").Superscript();
                        t.Span(" and H");
                        t.Span("2").Subscript();
                        t.Span("O.");
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("FontSize and FontColor: ");
                        t.Span("14pt orange").FontSize(14).FontColor(Colors.Orange.L700);
                        t.Span(", ");
                        t.Span("18pt bold blue").FontSize(18).Bold().FontColor(Colors.Blue.L700);
                        t.Span(". FontFamily: ");
                        t.Span("Courier New monospace").FontFamily("Courier New");
                        t.Span(".");
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("Highlight(HighlightColor): ");
                        t.Span(" Yellow ").Highlight(HighlightColor.Yellow);
                        t.Span(" Green ").Highlight(HighlightColor.Green);
                        t.Span(" Cyan ").Highlight(HighlightColor.Cyan);
                        t.Span(" Red ").Highlight(HighlightColor.Red);
                        t.Span(" DarkBlue ").Highlight(HighlightColor.DarkBlue).FontColor(Colors.White.Default);
                        t.Span(".");
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("Hex colors accept a leading '#' and 3-digit CSS shorthand: ");
                        t.Span("\"D32\" expands to \"DD3322\"").FontColor("D32");
                        t.Span(".");
                    });

                    col.Item().H2("3.2 Alignment & Justification");
                    col.Item().Text("AlignLeft() - the default paragraph alignment.");
                    col.Item().Text("AlignCenter() - centers this paragraph.", t => t.AlignCenter());
                    col.Item().Text("AlignRight() - right-aligns this paragraph.", t => t.AlignRight());
                    col.Item().Text("Justify() stretches every line except the last to fill the line width, which is " +
                                     "useful for body copy in reports, articles, and books where straight margins matter.",
                        t => t.Justify());

                    col.Item().H2("3.3 Spacing & Line Height");
                    col.Item().Text("LineHeight(1.5) increases the space between lines within this paragraph, while " +
                                     "SpacingBefore/SpacingAfter control the space above and below the paragraph as a whole - " +
                                     "useful for improving readability of long-form text.",
                        t => t.LineHeight(1.5f).SpacingBefore(6).SpacingAfter(6));

                    col.Item().H2("3.4 Indentation");
                    col.Item().Text("LeftIndent/RightIndent/FirstLineIndent: this paragraph is indented from both " +
                                     "margins and its first line is indented further, like a traditional book paragraph.",
                        t => t.LeftIndent(18).RightIndent(18).FirstLineIndent(18));
                    col.Item().Text(t =>
                    {
                        t.Span("HangingIndent: ").Bold();
                        t.Span("TerraFluent.Docx.Reporting. ");
                        t.Span("A zero-dependency .NET library for generating Word documents, with the first line " +
                               "flush left and following lines indented - the layout used for bibliographies and " +
                               "definition lists.");
                        t.LeftIndent(36).HangingIndent(36);
                    });

                    col.Item().H2("3.5 Paragraph Borders & Shading");
                    col.Item().Text(t =>
                    {
                        t.Span("Shading + Border (all sides): ").Bold();
                        t.Span("a fully boxed and shaded callout, ideal for executive summaries.");
                        t.Shading(Colors.Blue.L100).Border(1.25f, Colors.Blue.L700, 5).LeftIndent(8).RightIndent(8);
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("BorderLeft only: ").Bold();
                        t.Span("a single accent border on the left, ideal for warnings or notes.");
                        t.Shading(Colors.Orange.L100).BorderLeft(4, Colors.Orange.L700, 6).LeftIndent(12);
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("BorderTop + BorderBottom: ");
                        t.Span("horizontal rules above and below a quoted passage.", s => s.Italic());
                        t.BorderTop(0.75f, Colors.Grey.L400, 4).BorderBottom(0.75f, Colors.Grey.L400, 4).LeftIndent(18).RightIndent(18);
                    });

                    col.Item().H2("3.6 Tab Stops");
                    col.Item().Text(t =>
                    {
                        t.TabStop(260, TabStopAlignment.Right);
                        t.Span("Right-aligned tab stop");
                        t.Tab();
                        t.Span("100%").Bold();
                    });
                    col.Item().Text(t =>
                    {
                        t.TabStop(180, TabStopAlignment.Center);
                        t.TabStop(360, TabStopAlignment.Decimal);
                        t.Span("Center & decimal tabs");
                        t.Tab();
                        t.Span("MID");
                        t.Tab();
                        t.Span("1234.50");
                    });

                    col.Item().H2("3.7 Links, Bookmarks & Cross-References");
                    col.Item().Text(t =>
                    {
                        t.Span("Inline hyperlinks can be mixed into normal text, such as ");
                        t.Hyperlink("the TerraFluent.Docx.Reporting repository", "https://example.com/terrafluent-docx-reporting");
                        t.Span(".");
                    });
                    col.Item().Hyperlink("Standalone hyperlink paragraph via IContainer.Hyperlink", "https://example.com/docs");
                    col.Item().Text(t =>
                    {
                        t.Span("CrossReference resolves to the section's heading text and updates if the heading " +
                               "changes: see ");
                        t.CrossReference("sec-charts", "Section 7 (Charts)");
                        t.Span(" and ");
                        t.CrossReference("sec-tables", "Section 6 (Tables)");
                        t.Span(".");
                    });

                    col.Item().H2("3.8 Footnotes, Endnotes & Page Fields");
                    col.Item().Text(t =>
                    {
                        t.Span("Footnote() appends a real Word footnote");
                        t.Footnote("This is a footnote: it appears at the bottom of the page and is managed by Word's footnote part.");
                        t.Span(", and Endnote() appends a document-level endnote");
                        t.Endnote("This is an endnote: it appears in the document's endnotes section.");
                        t.Span(".");
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("CurrentPageNumber() and TotalPages() render as live Word fields: you are on page ");
                        t.CurrentPageNumber();
                        t.Span(" of ");
                        t.TotalPages();
                        t.Span(".");
                    });

                    col.Item().H2("3.9 Pagination Controls");
                    col.Item().Text("KeepWithNext() keeps this paragraph on the same page as the one that follows it - " +
                                     "useful for headings or lead-in sentences.", t => t.KeepWithNext());
                    col.Item().Text("KeepLinesTogether() prevents this paragraph from splitting across a page break, " +
                                     "and PageBreakBefore() (not used here) forces a paragraph to start a new page.",
                        t => t.KeepLinesTogether());
                });
            });

            // ----------------------------------------------------------------
            // Page 3 - Lists, in a two-column section layout
            // ----------------------------------------------------------------
            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.Columns(2, spacingPoints: 24, separatorLine: true);
                page.DefaultTextStyle(t => t.FontFamily("Aptos").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(4));
                AddStandardHeaderFooter(page, "4. Lists");

                page.Content().Column(col =>
                {
                    col.Spacing(6);

                    col.Item().Bookmark("sec-lists", "4. Lists", t => t.Style("Heading1"));
                    col.Item().Text("This section also demonstrates Columns(2, spacingPoints, separatorLine): the page " +
                                     "is split into two Word columns with a vertical separator line, and content " +
                                     "flows from the first column into the second.", t => t.Justify());

                    col.Item().H2("4.1 Bullet Lists");
                    col.Item().Text("BulletList supports custom markers per level via Marker(marker, level, fontFamily).");
                    col.Item().BulletList(list =>
                    {
                        list.Marker("•");
                        list.Marker("–", level: 1);
                        list.Item("First-level bullet item.");
                        list.Item("Second first-level item.");
                        list.Item("Nested second-level item.", level: 1);
                        list.Item("Another nested item.", level: 1);
                        list.Item(t =>
                        {
                            t.Span("Items can mix styled spans, like ");
                            t.Span("bold").Bold();
                            t.Span(" or ");
                            t.Span("italic").Italic();
                            t.Span(" text.");
                        });
                    });

                    col.Item().H2("4.2 Numbered Lists");
                    col.Item().Text("NumberedList supports custom number formats per level, such as \"%1.\" and \"%2)\".");
                    col.Item().NumberedList(list =>
                    {
                        list.Marker("%1.", level: 0);
                        list.Marker("%2)", level: 1);
                        list.Item("First numbered step.");
                        list.Item("Sub-step using the level-1 format.", level: 1);
                        list.Item("Another sub-step.", level: 1);
                        list.Item("Second top-level step.");
                        list.Item("Third top-level step.");
                    });

                    col.Item().Text(t =>
                    {
                        t.Span("Policy Note: ").Bold();
                        t.Span("Column layouts can still contain shaded callouts, borders, and lists side by side " +
                               "with normal flowing text.");
                        t.Style("Callout");
                    });
                });
            });

            // ----------------------------------------------------------------
            // Page 4 - Images
            // ----------------------------------------------------------------
            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.DefaultTextStyle(t => t.FontFamily("Aptos").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(4));
                AddStandardHeaderFooter(page, "5. Images");

                page.Content().Column(col =>
                {
                    col.Spacing(6);

                    col.Item().Bookmark("sec-images", "5. Images", t => t.Style("Heading1"));
                    col.Item().Text("IContainer.Image has overloads for a file path, a byte array, or a Stream. " +
                                     "IImageDescriptor then controls sizing, alignment, wrapping, position, and decoration.");

                    col.Item().H2("5.1 Sizing, Alt Text & Captions");
                    col.Item().Text("Width/Height/MaxWidth control sizing (aspect ratio is preserved when only one " +
                                     "dimension is given); AltText sets accessibility text; Caption adds an " +
                                     "auto-numbered \"Figure N.\" caption below the image.");
                    col.Item().Row(row =>
                    {
                        row.Spacing(10);
                        row.RelativeItem().Image(imagePath, img => img
                            .Width(80)
                            .AltText("Gradient image sized by width, from a file path")
                            .AlignCenter()
                            .Caption("From a file path, sized with Width(80)."));

                        row.RelativeItem().Image(imageBytes, "gradient-bytes.png", img => img
                            .Height(80)
                            .MaxWidth(110)
                            .AltText("Gradient image sized by height, from a byte array")
                            .AlignCenter()
                            .Caption("From a byte[], sized with Height(80) and MaxWidth(110)."));

                        using var imageStream = new MemoryStream(imageBytes);
                        row.RelativeItem().Image(imageStream, "gradient-stream.png", img => img
                            .Width(80)
                            .AltText("Gradient image from a stream")
                            .AlignCenter()
                            .Caption("From a Stream, sized with Width(80)."));
                    });

                    col.Item().H2("5.2 Wrapping & Floating");
                    col.Item().Text("WrapInline() is the default: the image sits inline with text like a large " +
                                     "character. WrapSquare/WrapTight let text wrap around a floated image; " +
                                     "FloatLeft/FloatRight/FloatCenter position it. WrapTopBottom keeps the image on " +
                                     "its own horizontal band; BehindText and InFrontOfText layer it with the text.");

                    col.Item().Image(imagePath, img => img
                        .Width(90)
                        .WrapSquare(8)
                        .FloatRight(8)
                        .AltText("Image floated right with square text wrapping"));
                    col.Item().Text("WrapSquare(marginPoints) combined with FloatRight(marginPoints) floats the " +
                                     "image to the right margin while this paragraph's text wraps around its " +
                                     "bounding box on the left and below. WrapTight instead follows the image's " +
                                     "contour more tightly (both render the same square wrap for this rectangular " +
                                     "sample image, but differ for non-rectangular pictures). This sentence is " +
                                     "intentionally long enough to flow around the floated image and demonstrate " +
                                     "the wrapping behaviour across multiple lines of text.", t => t.Justify());

                    col.Item().Image(imagePath, img => img
                        .Width(420)
                        .WrapTopBottom()
                        .AltText("Full-width banner image with top/bottom text wrapping")
                        .Caption("WrapTopBottom() defaults to horizontal-center alignment for banner-style images."));

                    col.Item().Text("BehindText() places the image on its own layer behind the paragraph text; " +
                                     "InFrontOfText() places it above the text. Both are commonly used for watermark-style " +
                                     "or decorative artwork and accept the same alignment, position, and margin controls " +
                                     "as the other wrap modes.");

                    col.Item().H2("5.3 Absolute Positioning");
                    col.Item().Text("Position(x, y) places a floating image relative to the page margins/paragraph; " +
                                     "PositionFromPage(x, y) measures from the physical page edges instead.");

                    col.Item().H2("5.4 Borders, Rounding & Cropping");
                    col.Item().Row(row =>
                    {
                        row.Spacing(10);
                        row.RelativeItem().Image(imagePath, img => img
                            .Width(80)
                            .Border(1.5f, Colors.Grey.L700)
                            .AlignCenter()
                            .Caption("Border(widthPoints, hexColor)."));

                        row.RelativeItem().Image(imagePath, img => img
                            .Width(80)
                            .Rounded()
                            .Border(1.5f, Colors.Blue.L700)
                            .AlignCenter()
                            .Caption("Rounded() combined with Border()."));

                        row.RelativeItem().Image(imagePath, img => img
                            .Width(80)
                            .Crop(10, 10, 10, 10)
                            .Border(1.5f, Colors.Green.L700)
                            .AlignCenter()
                            .Caption("Crop(left%, top%, right%, bottom%)."));
                    });
                });
            });

            // ----------------------------------------------------------------
            // Page 5 - Tables
            // ----------------------------------------------------------------
            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.DefaultTextStyle(t => t.FontFamily("Aptos").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(4));
                AddStandardHeaderFooter(page, "6. Tables");

                page.Content().Column(col =>
                {
                    col.Spacing(6);

                    col.Item().Bookmark("sec-tables", "6. Tables", t => t.Style("Heading1"));
                    col.Item().Text("ITableDescriptor configures column sizing, borders, padding, and banding. " +
                                     "ITableRowDescriptor adds rows, and ITableCellDescriptor (which extends " +
                                     "IContainer) configures each cell's content and appearance.");

                    col.Item().H2("6.1 Column Sizing");
                    col.Item().Text(t =>
                    {
                        t.Span("RelativeColumn(size)").Bold();
                        t.Span(" distributes remaining width proportionally; ");
                        t.Span("ConstantColumn(points)").Bold();
                        t.Span(" reserves an exact width. The table below mixes one 100pt constant column with two " +
                               "equal relative columns - the remaining page width (after the constant column) is " +
                               "split evenly between them.");
                    });
                    col.Item().Table(table =>
                    {
                        table.Style("Reference Table");
                        table.ColumnsDefinition(cols =>
                        {
                            cols.ConstantColumn(100);
                            cols.RelativeColumn();
                            cols.RelativeColumn();
                        });
                        table.HeaderRow(row =>
                        {
                            HeaderCell(row, "Fixed (100pt)");
                            HeaderCell(row, "Relative (1x)");
                            HeaderCell(row, "Relative (1x)");
                        });
                        table.Row(row =>
                        {
                            row.Cell().Text("ConstantColumn(100)");
                            row.Cell().Text("RelativeColumn()");
                            row.Cell().Text("RelativeColumn()");
                        });
                    });

                    col.Item().H2("6.2 Table Formatting & Styles");
                    col.Item().Text(".Style(\"Reference Table\") applies the named table style defined in section 1, " +
                                     "and WidthPercent(80).AlignCenter() sizes and centers the table on the page.");
                    col.Item().Table(table =>
                    {
                        table.Style("Reference Table");
                        table.WidthPercent(80).AlignCenter();
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(2);
                            cols.RelativeColumn(1);
                            cols.RelativeColumn(1);
                        });
                        table.HeaderRow(row =>
                        {
                            HeaderCell(row, "Setting");
                            HeaderCell(row, "Header Row");
                            HeaderCell(row, "Data Row");
                        });
                        table.Row(row =>
                        {
                            row.Cell().Text("HeaderBackground / AlternateRowBackground");
                            row.Cell().Text("Blue.L700", t => t.AlignRight());
                            row.Cell().Text("Grey.L100 (odd rows)", t => t.AlignRight());
                        });
                        table.Row(row =>
                        {
                            row.Cell().Text("HeaderRowMinHeight / RowMinHeight");
                            row.Cell().Text("22pt", t => t.AlignRight());
                            row.Cell().Text("20pt", t => t.AlignRight());
                        });
                        table.Row(row =>
                        {
                            row.Cell().Text("CellPadding(vertical, horizontal)");
                            row.Cell(2).Text("4pt, 6pt", t => t.AlignRight());
                        });
                    });

                    col.Item().H2("6.3 Cell Borders");
                    col.Item().Text("ITableCellDescriptor.Border applies all four sides at once; BorderTop/Right/Bottom/Left " +
                                     "set a single side independently.");
                    col.Item().Table(table =>
                    {
                        table.Border(0, Colors.White.Default);
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn();
                            cols.RelativeColumn();
                            cols.RelativeColumn();
                            cols.RelativeColumn();
                            cols.RelativeColumn();
                        });
                        table.Row(row =>
                        {
                            row.Cell().Padding(6).BorderTop(2, Colors.Blue.L700).Text("BorderTop", t => t.AlignCenter());
                            row.Cell().Padding(6).BorderRight(2, Colors.Blue.L700).Text("BorderRight", t => t.AlignCenter());
                            row.Cell().Padding(6).BorderBottom(2, Colors.Blue.L700).Text("BorderBottom", t => t.AlignCenter());
                            row.Cell().Padding(6).BorderLeft(2, Colors.Blue.L700).Text("BorderLeft", t => t.AlignCenter());
                            row.Cell().Padding(6).Border(2, Colors.Green.L700).Text("Border (all)", t => t.AlignCenter());
                        });
                    });

                    col.Item().H2("6.4 Cell Layout: Spans, Merges, Alignment & Text Direction");
                    col.Item().Table(table =>
                    {
                        table.Border(0.75f, Colors.Grey.L300);
                        table.RowMinHeight(36);
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn();
                            cols.RelativeColumn();
                            cols.ConstantColumn(40);
                            cols.ConstantColumn(40);
                        });

                        table.Row(row =>
                        {
                            row.Cell().VerticalMergeStart().VerticalAlignMiddle().Background(Colors.Blue.L100)
                                .Text("VerticalMergeStart / VerticalMergeContinue", t => t.Bold());
                            row.Cell().VerticalAlignTop().Text("VerticalAlignTop()");
                            row.Cell().ColumnSpan(2).TextDirectionTopToBottom().VerticalAlignMiddle()
                                .Text("TextDirectionTopToBottom", t => t.AlignCenter());
                        });
                        table.Row(row =>
                        {
                            row.Cell().VerticalMergeContinue();
                            row.Cell().VerticalAlignMiddle().Text("VerticalAlignMiddle()");
                            row.Cell().ColumnSpan(2).VerticalAlignMiddle().TextDirectionBottomToTop()
                                .Text("TextDirectionBottomToTop", t => t.AlignCenter());
                        });
                        table.Row(row =>
                        {
                            row.Cell(2).Padding(10).VerticalAlignBottom().Background(Colors.Grey.L100)
                                .Text("ColumnSpan(2) merges this cell across two columns; Padding(10) widens its margins; VerticalAlignBottom() aligns this text to the bottom of the cell.");
                            row.Cell(2).TextDirectionLeftToRight().VerticalAlignMiddle()
                                .Text("TextDirectionLeftToRight() (default)", t => t.AlignCenter());
                        });
                    });
                });
            });

            // ----------------------------------------------------------------
            // Page 6 - Charts and Layout Primitives (landscape)
            // ----------------------------------------------------------------
            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Landscape();
                page.Margin(Unit.Centimetre(1.8f));
                page.DefaultTextStyle(t => t.FontFamily("Aptos").FontSize(10).FontColor(Colors.Grey.L900).SpacingAfter(3));
                AddStandardHeaderFooter(page, "7-8. Charts & Layout");

                page.Content().Column(col =>
                {
                    col.Spacing(6);

                    col.Item().Bookmark("sec-charts", "7. Charts", t => t.Style("Heading1"));
                    col.Item().Text("IChartDescriptor.Series(configure) adds a data series; Series(name, configure) " +
                                     "additionally names the series for the legend. All series in one chart must use " +
                                     "the same point type (Bar/Line/Pie/Doughnut); pie and doughnut charts support only " +
                                     "a single series.");

                    col.Item().Row(row =>
                    {
                        row.Spacing(10);

                        row.RelativeItem().Column(c =>
                        {
                            c.Item().H2("7.1 Bar Chart (single series)");
                            c.Item().Chart(chart => chart
                                .Title("Quarterly Sales")
                                .Series(s => s
                                    .Color(Colors.Blue.L700)
                                    .Bar("Q1", 4.1)
                                    .Bar("Q2", 4.3)
                                    .Bar("Q3", 4.8)
                                    .Bar("Q4", 5.2)));
                        });

                        row.RelativeItem().Column(c =>
                        {
                            c.Item().H2("7.2 Bar Chart (multi-series)");
                            c.Item().Chart(chart => chart
                                .Title("Actual vs Target")
                                .Series("Actual", s => s
                                    .Color(Colors.Blue.L700)
                                    .Bar("Q1", 4.1)
                                    .Bar("Q2", 4.3)
                                    .Bar("Q3", 4.8)
                                    .Bar("Q4", 5.2))
                                .Series("Target", s => s
                                    .Color(Colors.Grey.L400)
                                    .Bar("Q1", 4.0)
                                    .Bar("Q2", 4.2)
                                    .Bar("Q3", 4.5)
                                    .Bar("Q4", 5.0)));
                        });

                        row.RelativeItem().Column(c =>
                        {
                            c.Item().H2("7.3 Line Chart (multi-series)");
                            c.Item().Chart(chart => chart
                                .Title("Year-over-Year Trend")
                                .Series("2025", s => s
                                    .Color(Colors.Grey.L500)
                                    .Line("Q1", 3.6).Line("Q2", 3.8).Line("Q3", 4.0).Line("Q4", 4.3))
                                .Series("2026", s => s
                                    .Color(Colors.Green.L700)
                                    .Line("Q1", 4.1).Line("Q2", 4.3).Line("Q3", 4.8).Line("Q4", 5.2)));
                        });
                    });

                    col.Item().Row(row =>
                    {
                        row.Spacing(10);

                        row.RelativeItem().Column(c =>
                        {
                            c.Item().H2("7.4 Pie Chart");
                            c.Item().Chart(chart => chart
                                .Title("Revenue By Practice")
                                .Series(s => s
                                    .Pie("Automation", 7.4)
                                    .Pie("Consulting", 4.8)
                                    .Pie("Compliance", 4.1)
                                    .Pie("Managed Delivery", 2.1)));
                        });

                        row.RelativeItem().Column(c =>
                        {
                            c.Item().H2("7.5 Doughnut Chart");
                            c.Item().Chart(chart => chart
                                .Title("Customer Segments")
                                .Series(s => s
                                    .Color(Colors.Orange.L700)
                                    .Doughnut("Enterprise", 52)
                                    .Doughnut("Mid-Market", 31)
                                    .Doughnut("Small Business", 17)));
                        });

                        row.RelativeItem().Column(c =>
                        {
                            c.Item().H2("8.1 Row & Column Sizing");
                            c.Item().Text("RelativeItem(size), AutoItem(), and ConstantItem(points) control how a " +
                                           "Row divides its width.", t => t.SpacingAfter(4));
                            c.Item().Row(r =>
                            {
                                r.Spacing(4);
                                r.ConstantItem(60).Text("Constant 60pt", t => t.FontSize(8.5f).Shading(Colors.Grey.L100));
                                r.AutoItem().Text("Auto", t => t.FontSize(8.5f).Shading(Colors.Blue.L100));
                                r.RelativeItem(2).Text("Relative 2x", t => t.FontSize(8.5f).Shading(Colors.Green.L100));
                                r.RelativeItem(1).Text("Relative 1x", t => t.FontSize(8.5f).Shading(Colors.Orange.L100));
                            });

                            c.Item().Bookmark("sec-layout", "8. Layout, Components & Templates", t => t.Style("Heading2"));
                            c.Item().Component(new CalloutComponent(
                                "8.2 Reusable Components",
                                "IContainer.Component(IComponent) composes a reusable fragment - this callout is " +
                                "rendered by a small IComponent implementation, reusing the \"Callout\" paragraph style " +
                                "from section 1."));

                            c.Item().Text(t =>
                            {
                                t.Span("8.3 Lines & Page Breaks: ").Bold();
                                t.Span("IContainer.Line() draws a horizontal rule (see below), and PageBreak() starts " +
                                       "a new page.");
                            });
                            c.Item().Line();

                            c.Item().Text(t =>
                            {
                                t.Span("8.4 Templates: ").Bold();
                                t.Span("DocxTemplate.Open(path).Replace(\"Key\", \"Value\").SaveAs(outputPath) performs " +
                                       "{{Key}} placeholder substitution on an existing .docx - see the separate " +
                                       "template-replacement sample.");
                            });

                            c.Item().Text(t =>
                            {
                                t.Span("This sentence refers back to ");
                                t.CrossReference("sec-typography", "Section 3 (Text & Typography)");
                                t.Span(" using CrossReference, completing the cross-reference pair started there.");
                            });
                        });
                    });
                });
            });
        }).PublishDocx(documentPath);

        return documentPath;
    }

    private static void AddStandardHeaderFooter(IPageDescriptor page, string sectionLabel)
    {
        page.Header().Row(row =>
        {
            row.RelativeItem().Text("TerraFluent.Docx.Reporting API Reference", t => t.Bold().FontColor(Colors.Blue.L800).FontSize(9));
            row.AutoItem().Text(sectionLabel, t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignRight());
        }).Line();

        page.Footer().Line().Row(row =>
        {
            row.RelativeItem().Text("TerraFluent.Docx.Reporting", t => t.FontSize(8.5f).FontColor(Colors.Grey.L600));
            row.AutoItem().Text(t =>
            {
                t.Span("Page ").FontSize(8.5f).FontColor(Colors.Grey.L600);
                t.CurrentPageNumber(s => s.FontSize(8.5f).FontColor(Colors.Grey.L600));
                t.Span(" of ").FontSize(8.5f).FontColor(Colors.Grey.L600);
                t.TotalPages(s => s.FontSize(8.5f).FontColor(Colors.Grey.L600));
                t.AlignRight();
            });
        });
    }

    private static void HeaderCell(ITableRowDescriptor row, string text) =>
        row.Cell().Text(text, t => t.Bold().FontColor(Colors.White.Default));

    private sealed class CalloutComponent(string label, string message) : IComponent
    {
        public void Compose(IContainer container)
        {
            container.Text(t =>
            {
                t.Span($"{label}: ").Bold();
                t.Span(message);
                t.Style("Callout");
            });
        }
    }
}
