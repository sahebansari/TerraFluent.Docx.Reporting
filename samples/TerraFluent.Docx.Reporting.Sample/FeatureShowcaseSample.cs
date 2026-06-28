using TerraFluent.Docx.Reporting;

internal static class FeatureShowcaseSample
{
    public static string Generate(string outputDirectory, string imagePath)
    {
        var documentPath = Path.Combine(outputDirectory, "sample-word-document.docx");

        Document.Create(container =>
        {
            container
                .Theme(theme => theme
                    .DefaultFont("Aptos", 10.5f)
                    .DefaultTextColor(Colors.Grey.L900)
                    .AccentColor(Colors.Blue.L700)
                    .HyperlinkColor(Colors.Red.L700)
                    .TableBorder(0.75f, Colors.Grey.L300)
                    .TableCellPadding(4, 6)
                    .TableHeaderBackground(Colors.Blue.L700)
                    .TableAlternateRowBackground(Colors.Grey.L100)
                    .TableHeaderRowMinHeight(22)
                    .TableRowMinHeight(20))
                .MetadataTitle("TerraFluent.Docx.Reporting Sample")
                .MetadataAuthor("TerraFluent.Docx.Reporting Author")
                .MetadataSubject("Demonstration - images, header, footer, page numbers");

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.54f));
                page.PageNumberStart(3);
                page.DefaultTextStyle(t => t.FontFamily("Calibri").FontSize(11).FontColor(Colors.Grey.L900).SpacingAfter(4));

                page.FirstPageHeader().Column(hdr =>
                {
                    hdr.Item().Text("TerraFluent.Docx.Reporting Sample Document - First Page", t => t.Bold().FontColor(Colors.Blue.L700).AlignCenter());
                    hdr.Item().Line();
                });

                page.EvenPageHeader().Column(hdr =>
                {
                    hdr.Item().Row(row =>
                    {
                        row.RelativeItem().Text("TerraFluent.Docx.Reporting Sample Document", t => t.Bold().FontColor(Colors.Grey.L700));
                        row.AutoItem().Text("Even page", t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignRight());
                    });
                    hdr.Item().Line();
                });

                page.Header().Column(hdr =>
                {
                    hdr.Item().Row(row =>
                    {
                        row.RelativeItem().Text("TerraFluent.Docx.Reporting Sample Document", t => t.Bold().FontColor(Colors.Blue.L700));
                        row.AutoItem().Text(t => t.CurrentPageNumber().AlignRight());
                    });
                    hdr.Item().Line();
                });

                page.Footer().Column(ftr =>
                {
                    ftr.Item().Line();
                    ftr.Item().Text(t =>
                    {
                        t.Span("Page ").FontSize(9).FontColor(Colors.Grey.L600);
                        t.CurrentPageNumber(s => s.FontSize(9).FontColor(Colors.Grey.L600));
                        t.Span(" of ").FontSize(9).FontColor(Colors.Grey.L600);
                        t.TotalPages(s => s.FontSize(9).FontColor(Colors.Grey.L600));
                        t.AlignCenter();
                    });
                });

                page.FirstPageFooter().Text("First page footer uses a separate Word header/footer part.", t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignCenter());
                page.EvenPageFooter().Text("Even page footer", t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignCenter());

                page.Content().Column(col =>
                {
                    col.Spacing(8);

                    col.Item().H1("TerraFluent.Docx.Reporting");
                    col.Item().H2("Zero-dependency .docx generation for .NET");

                    col.Item().Text("This document demonstrates header/footer rendering, " +
                                    "image embedding, and page number fields - all generated " +
                                    "as valid OOXML with no external dependencies.", t => t.SpacingAfter(6));

                    col.Item().H3("Image Embedding");
                    col.Item().Text("The image below is embedded directly from a file. " +
                                    "Width is specified in points; height is computed from the aspect ratio.");
                    col.Item().Image(imagePath, img => img
                        .Width(150)
                        .MaxWidth(180)
                        .AltText("Blue and orange generated gradient square")
                        .AlignCenter()
                        .Caption("Figure 1. File-backed image with centered alignment and alt text."));

                    var imageBytes = File.ReadAllBytes(imagePath);
                    col.Item().Text("The same image can also be embedded from byte arrays or streams.");
                    col.Item().Row(row =>
                    {
                        row.Spacing(8);
                        row.RelativeItem().Image(imageBytes, "byte-array-gradient.png", img => img
                            .Width(70)
                            .AltText("Gradient image embedded from byte array")
                            .AlignCenter()
                            .Caption("Byte array"));

                        using var imageStream = new MemoryStream(imageBytes);
                        row.RelativeItem().Image(imageStream, "stream-gradient.png", img => img
                            .Height(70)
                            .MaxWidth(80)
                            .AltText("Gradient image embedded from stream")
                            .AlignCenter()
                            .Caption("Stream"));
                    });

                    col.Item().H3("Text Styling");
                    col.Item().Text(t =>
                    {
                        t.Span("Bold, ").Bold();
                        t.Span("italic, ").Italic();
                        t.Span("underlined, ").Underline();
                        t.Span("strikethrough. ").Strikethrough();
                        t.Span("Custom ").FontColor(Colors.Orange.L700).FontSize(14);
                        t.Span("colours").FontColor(Colors.Blue.L700).Bold();
                        t.Span(" and sizes.");
                    });

                    col.Item().H3("Superscript, Subscript & Highlights");
                    col.Item().Text(t =>
                    {
                        t.Span("Scientific notation uses Superscript() and Subscript(): E=mc");
                        t.Span("2").Superscript();
                        t.Span(" and H");
                        t.Span("2").Subscript();
                        t.Span("O.");
                    });
                    col.Item().Text(t =>
                    {
                        t.Span("E=mc");
                        t.Span("2").Superscript();
                        t.Span(" is highlighted: ");
                        t.Span("important").Highlight(HighlightColor.Yellow);
                    });

                    col.Item().H3("Hyperlinks");
                    col.Item().Text(t =>
                    {
                        t.Span("Inline links can be mixed into normal text, such as ");
                        t.Hyperlink("the TerraFluent.Docx.Reporting project page", "https://example.com/terrafluent-docx-reporting");
                        t.Span(".");
                    });
                    col.Item().Hyperlink("Standalone hyperlink paragraph", "https://example.com/docs", t => t.FontSize(11));

                    col.Item().H3("Alignment");
                    col.Item().Text("Left aligned (default).");
                    col.Item().Text("Centered.", t => t.AlignCenter());
                    col.Item().Text("Right aligned.", t => t.AlignRight());
                    col.Item().Text("Justified - a longer sentence to demonstrate how " +
                                    "justified alignment fills the full page width.", t => t.Justify());

                    col.Item().H3("Paragraph Controls");
                    col.Item().Text("This paragraph has left and right indentation plus a first-line indent. " +
                                    "It also keeps with the following paragraph, which is useful for headings, clauses, " +
                                    "and short explanatory blocks in professional documents.",
                        t => t.LeftIndent(18).RightIndent(12).FirstLineIndent(12).SpacingBefore(4).SpacingAfter(4).KeepWithNext());
                    col.Item().Text("This follow-up paragraph is kept with the previous one and uses tighter spacing.",
                        t => t.LeftIndent(18).SpacingAfter(2).KeepLinesTogether());

                    col.Item().H3("Lists");
                    col.Item().Text("The items below use real Word numbering definitions.");
                    col.Item().BulletList(list =>
                    {
                        list.Marker("✓", fontFamily: "Segoe UI Symbol");
                        list.Marker("→", level: 1, fontFamily: "Segoe UI Symbol");
                        list.Item("Bullet list item with inherited default text styling.");
                        list.Item("Nested bullet item with a custom marker.", level: 1);
                        list.Item(t =>
                        {
                            t.Span("Styled bullet item with ");
                            t.Span("bold emphasis").Bold();
                            t.Span(".");
                        });
                    });
                    col.Item().NumberedList(list =>
                    {
                        list.Marker("%1)", level: 0);
                        list.Marker("%2)", level: 1);
                        list.Item("First numbered step.");
                        list.Item("Nested numbered detail.", level: 1);
                        list.Item("Second numbered step.");
                    });

                    col.Item().Line();

                    col.Item().H3("Two-Column Row");
                    col.Item().Row(row =>
                    {
                        row.Spacing(6);
                        row.RelativeItem().Column(left =>
                        {
                            left.Item().H4("Left");
                            left.Item().Text("Content in the left column.");
                            left.Item().Image(imagePath, img => img.Width(80).AltText("Small generated gradient image"));
                        });
                        row.RelativeItem().Column(right =>
                        {
                            right.Item().H4("Right");
                            right.Item().Text("Content in the right column.");
                        });
                    });

                    col.Item().H3("Table");
                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(2);
                            cols.RelativeColumn(1);
                            cols.RelativeColumn(1);
                        });

                        table.HeaderRow(row =>
                        {
                            row.Cell().Text("Name", t => t.Bold().FontColor(Colors.White.Default));
                            row.Cell().Text("Role", t => t.Bold().FontColor(Colors.White.Default));
                            row.Cell().Text("Score", t => t.Bold().FontColor(Colors.White.Default));
                        });
                        table.Row(r => { r.Cell().Text("Alice"); r.Cell().Text("Engineer"); r.Cell().Text("98"); });
                        table.Row(r => { r.Cell().Text("Bob"); r.Cell().Text("Designer"); r.Cell().Text("87"); });
                        table.Row(r => { r.Cell().Text("Carol"); r.Cell().Text("Manager"); r.Cell().Text("92"); });
                    });

                    col.Item().PageBreak();

                    col.Item().H2("Page 2");
                    col.Item().Text("This content follows a page break. The header and footer defined " +
                                    "on the page section appear on every page. First-page and even-page " +
                                    "header/footer variants are also available.");
                    col.Item().H3("Heading Levels");
                    col.Item().H1("H1");
                    col.Item().H2("H2");
                    col.Item().H3("H3");
                    col.Item().H4("H4");
                    col.Item().H5("H5");
                    col.Item().H6("H6");
                });
            });
        }).PublishDocx(documentPath);

        return documentPath;
    }
}
