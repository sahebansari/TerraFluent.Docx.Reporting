using TerraFluent.Docx.Reporting;

internal static class LayoutFeaturesSample
{
    public static string Generate(string outputDirectory)
    {
        var documentPath = Path.Combine(outputDirectory, "sample-layout-features.docx");

        Document.Create(container =>
        {
            container
                .MetadataTitle("TerraFluent.Docx.Reporting Layout Features")
                .MetadataAuthor("TerraFluent.Docx.Reporting Sample")
                .MetadataSubject("Paragraph borders, paragraph shading, and page columns")
                .MetadataCreator("TerraFluent.Docx.Reporting Sample App")
                .Theme(theme => theme
                    .DefaultFont("Aptos", 10.5f)
                    .HeadingColor(Colors.Blue.L800)
                    .AccentColor(Colors.Blue.L700));

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.2f));
                page.Header().Text("Paragraph Borders And Shading", t => t.Bold().FontColor(Colors.Blue.L800).AlignCenter());
                page.Footer().Text(t =>
                {
                    t.Span("Page ");
                    t.CurrentPageNumber();
                    t.AlignCenter();
                });

                page.Content().Column(col =>
                {
                    col.Spacing(8);
                    col.Item().H1("Paragraph Callouts");
                    col.Item().Text("Paragraph borders and shading are useful for executive summaries, warning blocks, quotes, and policy notes.");

                    col.Item().Text(t =>
                    {
                        t.Span("Executive Summary: ").Bold();
                        t.Span("Revenue grew across every operating practice, while delivery cycle time improved by 18%. The leadership team should continue investing in reusable document workflows.");
                        t.Shading(Colors.Blue.L100)
                            .Border(1.25f, Colors.Blue.L700, 5)
                            .LeftIndent(10)
                            .RightIndent(10)
                            .SpacingBefore(8)
                            .SpacingAfter(8)
                            .KeepLinesTogether();
                    });

                    col.Item().Text(t =>
                    {
                        t.Span("Warning: ").Bold().FontColor(Colors.Orange.L900);
                        t.Span("Client-facing reports should be reviewed before release when generated from draft source data.");
                        t.Shading(Colors.Orange.L100)
                            .BorderLeft(4, Colors.Orange.L700, 5)
                            .LeftIndent(14)
                            .SpacingBefore(8)
                            .SpacingAfter(8)
                            .KeepLinesTogether();
                    });

                    col.Item().Text(t =>
                    {
                        t.Span("Document automation works best when design decisions are encoded once and reused everywhere.", s => s.Italic());
                        t.Shading(Colors.Grey.L100)
                            .BorderTop(0.75f, Colors.Grey.L400, 4)
                            .BorderBottom(0.75f, Colors.Grey.L400, 4)
                            .LeftIndent(18)
                            .RightIndent(18)
                            .SpacingBefore(10)
                            .SpacingAfter(10);
                    });
                });
            });

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(1.8f));
                page.Columns(2, spacingPoints: 24, separatorLine: true);
                page.Header().Text("Two-Column Newsletter Section", t => t.Bold().FontColor(Colors.Blue.L800).AlignCenter());
                page.Footer().Text(t =>
                {
                    t.Span("Page ");
                    t.CurrentPageNumber();
                    t.AlignCenter();
                });

                page.Content().Column(col =>
                {
                    col.Spacing(5);
                    col.Item().H1("Operations Brief");
                    col.Item().Text("This page uses section-level Word columns. Normal paragraphs flow from the first column into the second column automatically.");

                    for (int i = 1; i <= 8; i++)
                    {
                        col.Item().Text($"Update {i}: Teams standardized their report sections, reusable tables, and review checkpoints. The result is a clearer authoring workflow and more consistent client-ready documents.", t => t.Justify());
                    }

                    col.Item().Text(t =>
                    {
                        t.Span("Policy Note: ").Bold();
                        t.Span("Column layouts can still contain shaded callouts and bordered notes.");
                        t.Shading(Colors.Green.L100)
                            .Border(0.75f, Colors.Green.L700, 4)
                            .SpacingBefore(8)
                            .SpacingAfter(8);
                    });
                });
            });
        }).PublishDocx(documentPath);

        return documentPath;
    }
}
