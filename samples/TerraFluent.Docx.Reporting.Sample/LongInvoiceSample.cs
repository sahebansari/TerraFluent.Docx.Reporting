using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;

internal static class LongInvoiceSample
{
    public static string Generate(string outputDirectory, string logoPath)
    {
        var documentPath = Path.Combine(outputDirectory, "sample-long-invoice.docx");
        var issueDate = new DateTime(2026, 6, 7);
        var dueDate = issueDate.AddDays(30);
        var items = CreateLineItems();
        var subtotal = items.Sum(item => item.Quantity * item.Rate);
        var tax = Math.Round(subtotal * 0.085m, 2);
        var total = subtotal + tax;

        Document.Create(container =>
        {
            container
                .MetadataTitle("Sample Long Invoice")
                .MetadataAuthor("TerraFluent.Docx.Reporting Sample")
                .MetadataSubject("Multi-page invoice with repeating table header");

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(1.8f));
                page.DefaultTextStyle(t => t.FontFamily("Calibri").FontSize(10).FontColor(Colors.Grey.L900).SpacingAfter(2));

                InvoiceBranding.Header(page.Header(), logoPath, "LONG INVOICE", 15);
                InvoiceBranding.Footer(page.Footer(), logoPath, includePageNumbers: true);

                page.Content().Column(col =>
                {
                    col.Spacing(6);

                    col.Item().Row(row =>
                    {
                        row.Spacing(8);
                        row.RelativeItem().Column(left =>
                        {
                            left.Item().Text("Bill To", t => t.Bold().FontColor(Colors.Blue.L700));
                            left.Item().Text("Acme Industries");
                            left.Item().Text("42 Market Street");
                            left.Item().Text("San Francisco, CA 94105");
                            left.Item().Text("billing@acme.example");
                        });

                        row.RelativeItem().Column(right =>
                        {
                            right.Item().Text("Invoice Details", t => t.Bold().FontColor(Colors.Blue.L700).AlignRight());
                            right.Item().Text("Invoice #: INV-2026-0098", t => t.AlignRight());
                            right.Item().Text($"Issue date: {issueDate:MMM dd, yyyy}", t => t.AlignRight());
                            right.Item().Text($"Due date: {dueDate:MMM dd, yyyy}", t => t.AlignRight());
                            right.Item().Text("Terms: Net 30", t => t.AlignRight());
                        });
                    });

                    col.Item().Line();

                    col.Item().Text("Professional Services", t => t.Bold().FontSize(12).FontColor(Colors.Blue.L700));
                    col.Item().Table(table =>
                    {
                        table
                            .CellPadding(4, 6)
                            .HeaderBackground(Colors.Blue.L700)
                            .AlternateRowBackground(Colors.Grey.L100)
                            .HeaderRowMinHeight(22)
                            .RowMinHeight(20)
                            .Border(0.75f, Colors.Grey.L300);

                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(4);
                            cols.RelativeColumn(1);
                            cols.RelativeColumn(1);
                            cols.RelativeColumn(1);
                        });

                        table.HeaderRow(row =>
                        {
                            row.Cell().Text("Description", t => t.Bold().FontColor(Colors.White.Default));
                            row.Cell().Text("Qty", t => t.Bold().FontColor(Colors.White.Default).AlignRight());
                            row.Cell().Text("Rate", t => t.Bold().FontColor(Colors.White.Default).AlignRight());
                            row.Cell().Text("Amount", t => t.Bold().FontColor(Colors.White.Default).AlignRight());
                        });

                        foreach (var item in items)
                            AddLineItem(table, item);

                        AddSummaryRow(table, "Subtotal", subtotal, Colors.Grey.L100, bold: false);
                        AddSummaryRow(table, "Tax (8.5%)", tax, Colors.Grey.L100, bold: false);
                        AddSummaryRow(table, "Total Due", total, Colors.Blue.L700, bold: true, invertText: true);
                    });

                    col.Item().Text("Payment Notes", t => t.Bold().FontColor(Colors.Blue.L700));
                    col.Item().BulletList(list =>
                    {
                        list.Item("This sample intentionally includes enough invoice lines to continue the table onto the next page.");
                        list.Item("The item table header row is marked as a repeating Word table header.");
                    });

                    col.Item().Text("Remit payment to: Northwind Consulting, 100 Business Plaza, Seattle, WA 98101");
                });
            });
        }).PublishDocx(documentPath);

        return documentPath;
    }

    private static IReadOnlyList<InvoiceLineItem> CreateLineItems() =>
    [
        new("Discovery workshop and requirements mapping", 1, 1200m),
        new("Stakeholder interview synthesis", 3, 425m),
        new("Document automation prototype", 2, 950m),
        new("Invoice template data model", 1, 875m),
        new("OOXML table styling implementation", 2, 800m),
        new("Header and footer layout pass", 1, 540m),
        new("Multi-page table QA pass", 2, 375m),
        new("Accessibility and metadata review", 1, 460m),
        new("Sample document design polish", 2, 390m),
        new("Regression test authoring", 3, 325m),
        new("Field and page-number verification", 1, 350m),
        new("Invoice totals validation", 1, 300m),
        new("Package relationship inspection", 1, 280m),
        new("Word compatibility smoke test", 2, 260m),
        new("Alternate row shading review", 1, 240m),
        new("Table padding and row-height tuning", 2, 310m),
        new("Long-form sample content setup", 1, 295m),
        new("Client feedback incorporation", 2, 450m),
        new("Implementation handoff notes", 1, 520m),
        new("Documentation example cleanup", 2, 275m),
        new("Build pipeline sample validation", 1, 360m),
        new("Final review and delivery package", 1, 690m),
        new("Post-delivery support block", 2, 225m),
        new("Template reuse consultation", 1, 410m),
        new("Output folder automation update", 1, 250m),
        new("Invoice sample extension", 2, 330m),
        new("Table continuation testing", 2, 345m),
        new("Document repair prompt verification", 1, 380m),
        new("Sample cleanup and formatting pass", 1, 315m),
        new("Release readiness review", 1, 475m)
    ];

    private static void AddLineItem(ITableDescriptor table, InvoiceLineItem item)
    {
        var amount = item.Quantity * item.Rate;
        table.Row(row =>
        {
            row.Cell().VerticalAlignMiddle().Text(item.Description);
            row.Cell().VerticalAlignMiddle().Text(item.Quantity.ToString(), t => t.AlignRight());
            row.Cell().VerticalAlignMiddle().Text(item.Rate.ToString("$#,##0.00"), t => t.AlignRight());
            row.Cell().VerticalAlignMiddle().Text(amount.ToString("$#,##0.00"), t => t.AlignRight());
        });
    }

    private static void AddSummaryRow(ITableDescriptor table, string label, decimal amount, string backgroundColor, bool bold, bool invertText = false)
    {
        table.Row(row =>
        {
            row.Cell(3)
                .Background(backgroundColor)
                .Padding(5, 7)
                .VerticalAlignMiddle()
                .Text(label, t =>
                {
                    t.AlignRight();
                    if (bold) t.Bold().FontSize(12);
                    if (invertText) t.FontColor(Colors.White.Default);
                });
            row.Cell()
                .Background(backgroundColor)
                .Padding(5, 7)
                .VerticalAlignMiddle()
                .Text(amount.ToString("$#,##0.00"), t =>
                {
                    t.AlignRight();
                    if (bold) t.Bold().FontSize(12);
                    if (invertText) t.FontColor(Colors.White.Default);
                });
        });
    }

    private sealed record InvoiceLineItem(string Description, int Quantity, decimal Rate);
}
