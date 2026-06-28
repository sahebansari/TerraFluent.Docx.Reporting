using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;

internal static class InvoiceSample
{
    public static string Generate(string outputDirectory, string logoPath)
    {
        var documentPath = Path.Combine(outputDirectory, "sample-invoice.docx");
        var issueDate = new DateTime(2026, 6, 7);
        var dueDate = issueDate.AddDays(14);

        Document.Create(container =>
        {
            container
                .MetadataTitle("Sample Invoice")
                .MetadataAuthor("TerraFluent.Docx.Reporting Sample")
                .MetadataSubject("Invoice document generated with TerraFluent.Docx.Reporting");

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.DefaultTextStyle(t => t.FontFamily("Calibri").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(3));

                InvoiceBranding.Header(page.Header(), logoPath, "INVOICE", 16);
                InvoiceBranding.Footer(page.Footer(), logoPath, includePageNumbers: false);

                page.Content().Column(col =>
                {
                    col.Spacing(7);

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
                            right.Item().Text("Invoice #: INV-2026-0042", t => t.AlignRight());
                            right.Item().Text($"Issue date: {issueDate:MMM dd, yyyy}", t => t.AlignRight());
                            right.Item().Text($"Due date: {dueDate:MMM dd, yyyy}", t => t.AlignRight());
                            right.Item().Text("Terms: Net 14", t => t.AlignRight());
                        });
                    });

                    col.Item().Line();

                    col.Item().Table(table =>
                    {
                        table
                            .CellPadding(5, 7)
                            .HeaderBackground(Colors.Blue.L700)
                            .AlternateRowBackground(Colors.Grey.L100)
                            .HeaderRowMinHeight(24)
                            .RowMinHeight(22)
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

                        AddLineItem(table, "Discovery workshop and requirements mapping", 1, 1200m);
                        AddLineItem(table, "Document automation prototype", 2, 950m);
                        AddLineItem(table, "Implementation review and handoff", 1, 750m);
                        AddSummaryRow(table, "Subtotal", 3850m, Colors.Grey.L100, bold: false);
                        AddSummaryRow(table, "Tax (8.5%)", 327.25m, Colors.Grey.L100, bold: false);
                        AddSummaryRow(table, "Total Due", 4177.25m, Colors.Blue.L700, bold: true, invertText: true);
                    });

                    col.Item().Text("Payment Notes", t => t.Bold().FontColor(Colors.Blue.L700));
                    col.Item().BulletList(list =>
                    {
                        list.Item("Please include the invoice number with payment.");
                        list.Item("Late payments may be subject to a 1.5% monthly finance charge.");
                    });

                    col.Item().Text("Remit payment to: Northwind Consulting, 100 Business Plaza, Seattle, WA 98101");
                });
            });
        }).PublishDocx(documentPath);

        return documentPath;
    }

    private static void AddLineItem(ITableDescriptor table, string description, int quantity, decimal rate)
    {
        var amount = quantity * rate;
        table.Row(row =>
        {
            row.Cell().VerticalAlignMiddle().Text(description);
            row.Cell().VerticalAlignMiddle().Text(quantity.ToString(), t => t.AlignRight());
            row.Cell().VerticalAlignMiddle().Text(rate.ToString("$#,##0.00"), t => t.AlignRight());
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
}
