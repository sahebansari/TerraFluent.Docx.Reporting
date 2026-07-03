using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// A protected "review draft" market analysis showcasing the newer features: auto-numbered
/// figure/table captions with a table of figures and list of tables, chart legend placement,
/// axis titles, data labels and stacking, chart sizing/alignment, and comments-only document
/// protection so reviewers can annotate but not edit.
/// </summary>
internal static class MarketAnalysisSample
{
    public static string Generate(string outputDirectory, string imagePath)
    {
        var documentPath = Path.Combine(outputDirectory, "sample-market-analysis.docx");

        Document.Create(container =>
        {
            container
                .MetadataTitle("Market Analysis - Review Draft")
                .MetadataAuthor("TerraFluent.Docx.Reporting Sample")
                .MetadataSubject("Showcase of captions, tables of figures, chart options, and restricted editing");

            // Reviewers can only add comments; the password to lift the restriction is "review2026".
            container.RestrictEditing(DocumentProtection.CommentsOnly, "review2026");

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.DefaultTextStyle(t => t.FontFamily("Calibri").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(4));

                page.Header().Row(row =>
                {
                    row.RelativeItem().Text("Market Analysis 2026", t => t.Bold().FontColor(Colors.Blue.L700));
                    row.AutoItem().Text("REVIEW DRAFT", t => t.Bold().FontColor(Colors.Red.L700).AlignRight());
                });

                page.Footer().Text(t =>
                {
                    t.Span("Comments only - unprotect password: review2026   |   Page ");
                    t.CurrentPageNumber();
                    t.Span(" of ");
                    t.TotalPages();
                    t.AlignCenter();
                    t.FontSize(8.5f);
                    t.FontColor(Colors.Grey.L600);
                });

                page.Content().Column(col =>
                {
                    col.Spacing(10);

                    col.Item().H1("Market Analysis 2026");
                    col.Item().Text("This document is protected as comments-only: reviewers can add comments in Word " +
                                    "but cannot change the text. Every figure and table below carries an auto-numbered " +
                                    "caption, collected into the lists that follow. Right-click a list and choose " +
                                    "\"Update Field\" to refresh numbering and page references.");

                    col.Item().TableOfFigures();
                    col.Item().TableOfFigures("List of Tables", "Table");
                    col.Item().PageBreak();

                    col.Item().H2("Segment Performance");
                    col.Item().Text("Enterprise remains the largest contributor, with mid-market growing fastest. " +
                                    "The stacked chart shows each segment's share of total revenue per quarter.");
                    col.Item().Chart(chart => chart
                        .Title("Revenue by Segment")
                        .Width(430)
                        .Height(250)
                        .AlignCenter()
                        .Legend(ChartLegendPosition.Bottom)
                        .CategoryAxisTitle("Quarter")
                        .ValueAxisTitle("Revenue ($M)")
                        .DataLabels()
                        .Stacked()
                        .Series("Enterprise", s => s.Color(Colors.Blue.L700)
                            .Bar("Q1", 6.2).Bar("Q2", 6.5).Bar("Q3", 6.9).Bar("Q4", 7.4))
                        .Series("Mid-Market", s => s.Color(Colors.Green.L700)
                            .Bar("Q1", 3.1).Bar("Q2", 3.6).Bar("Q3", 4.2).Bar("Q4", 4.9))
                        .Series("Small Business", s => s.Color(Colors.Orange.L700)
                            .Bar("Q1", 1.4).Bar("Q2", 1.5).Bar("Q3", 1.6).Bar("Q4", 1.8)));

                    col.Item().Image(imagePath, img => img
                        .Width(140)
                        .AlignCenter()
                        .AltText("Brand artwork for the market analysis")
                        .FigureCaption("Market coverage artwork generated for this draft."));

                    col.Item().H2("Regional Results");
                    col.Item().Table(t =>
                    {
                        t.Caption("Revenue and growth by region, full year 2026.");
                        t.ColumnsDefinition(d => { d.RelativeColumn(2); d.RelativeColumn(1); d.RelativeColumn(1); });
                        t.CellPadding(4, 6);
                        t.Border(0.75f, Colors.Grey.L300);
                        t.HeaderBackground(Colors.Blue.L700);
                        t.HeaderRow(r =>
                        {
                            r.Cell().Text("Region", x => x.Bold().FontColor(Colors.White.Default));
                            r.Cell().Text("Revenue", x => x.Bold().FontColor(Colors.White.Default).AlignRight());
                            r.Cell().Text("Growth", x => x.Bold().FontColor(Colors.White.Default).AlignRight());
                        });
                        t.Row(r => { r.Cell().Text("North America"); r.Cell().Text("$18.4M", x => x.AlignRight()); r.Cell().Text("+12%", x => x.AlignRight()); });
                        t.Row(r => { r.Cell().Text("Europe"); r.Cell().Text("$11.7M", x => x.AlignRight()); r.Cell().Text("+9%", x => x.AlignRight()); });
                        t.Row(r => { r.Cell().Text("Asia-Pacific"); r.Cell().Text("$7.2M", x => x.AlignRight()); r.Cell().Text("+21%", x => x.AlignRight()); });
                    });

                    col.Item().H2("Customer Trend");
                    col.Item().Chart(chart => chart
                        .Title("Active Customers")
                        .Width(430)
                        .Height(230)
                        .AlignCenter()
                        .HideLegend()
                        .DataLabels()
                        .CategoryAxisTitle("Month")
                        .ValueAxisTitle("Customers")
                        .Series("Customers", s => s.Color(Colors.Green.L700)
                            .Line("Jan", 480).Line("Feb", 505).Line("Mar", 538)
                            .Line("Apr", 561).Line("May", 590).Line("Jun", 624)));

                    col.Item().Table(t =>
                    {
                        t.Caption("Reviewer sign-off checklist.");
                        t.ColumnsDefinition(d => { d.RelativeColumn(3); d.RelativeColumn(1); });
                        t.CellPadding(4, 6);
                        t.Border(0.75f, Colors.Grey.L300);
                        t.Row(r => { r.Cell().Text("Numbers reconciled with finance export"); r.Cell().Text("[ ]", x => x.AlignCenter()); });
                        t.Row(r => { r.Cell().Text("Regional commentary reviewed"); r.Cell().Text("[ ]", x => x.AlignCenter()); });
                        t.Row(r => { r.Cell().Text("Ready for distribution"); r.Cell().Text("[ ]", x => x.AlignCenter()); });
                    });
                });
            });
        }).PublishDocx(documentPath);

        return documentPath;
    }
}
