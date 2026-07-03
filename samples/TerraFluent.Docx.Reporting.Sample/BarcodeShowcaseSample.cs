using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;

internal static class BarcodeShowcaseSample
{
    private static readonly (string Sku, string Description, int Qty)[] Packages =
    [
        ("SKU-10042-BLK", "Wireless Mouse, Black", 24),
        ("SKU-10043-WHT", "Wireless Mouse, White", 12),
        ("SKU-20871-M", "Cotton T-Shirt, Medium", 48),
        ("SKU-20871-L", "Cotton T-Shirt, Large", 36),
        ("SKU-33190", "USB-C Charging Cable, 2m", 60),
    ];

    public static string Generate(string outputDirectory, string logoPath)
    {
        var documentPath = Path.Combine(outputDirectory, "sample-barcode-showcase.docx");
        var shipDate = new DateTime(2026, 7, 3);
        const string trackingNumber = "1Z999AA10123456784";

        Document.Create(container =>
        {
            container
                .MetadataTitle("Sample Shipping Manifest")
                .MetadataAuthor("TerraFluent.Docx.Reporting Sample")
                .MetadataSubject("Barcode feature showcase generated with TerraFluent.Docx.Reporting");

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.0f));
                page.DefaultTextStyle(t => t.FontFamily("Calibri").FontSize(10.5f).FontColor(Colors.Grey.L900).SpacingAfter(3));

                InvoiceBranding.Header(page.Header(), logoPath, "SHIPPING MANIFEST", 14);
                InvoiceBranding.Footer(page.Footer(), logoPath, includePageNumbers: true);

                page.Content().Column(col =>
                {
                    col.Spacing(8);

                    col.Item().H1("Shipping Manifest");
                    col.Item().Text("This document demonstrates the barcode fluent API: Code 128 barcodes rendered " +
                                    "as vector shapes, usable anywhere an image can go - page body, headers, footers, " +
                                    "columns, rows, and table cells.");

                    col.Item().Line();

                    col.Item().Row(row =>
                    {
                        row.Spacing(2);
                        row.RelativeItem().Column(left =>
                        {
                            left.Item().Text("Ship Date", t => t.Bold().FontColor(Colors.Blue.L700));
                            left.Item().Text($"{shipDate:MMM dd, yyyy}");
                            left.Item().Text("Carrier", t => t.Bold().FontColor(Colors.Blue.L700).SpacingBefore(6));
                            left.Item().Text("Northwind Freight & Logistics");
                        });

                        row.RelativeItem().Column(right =>
                        {
                            right.Item().Text("Tracking Number", t => t.Bold().FontColor(Colors.Blue.L700).AlignCenter());
                            right.Item().Barcode(trackingNumber, bc => bc
                                .Width(220)
                                .Height(50)
                                .AlignCenter()
                                .AltText($"Tracking barcode {trackingNumber}"));
                        });
                    });

                    col.Item().H2("Package Contents");
                    col.Item().Table(table =>
                    {
                        table
                            .CellPadding(5, 7)
                            .HeaderBackground(Colors.Blue.L700)
                            .AlternateRowBackground(Colors.Grey.L100)
                            .HeaderRowMinHeight(24)
                            .RowMinHeight(34)
                            .Border(0.75f, Colors.Grey.L300);

                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(2);
                            cols.RelativeColumn(3);
                            cols.RelativeColumn(1);
                            cols.RelativeColumn(2);
                        });

                        table.HeaderRow(row =>
                        {
                            row.Cell().Text("SKU", t => t.Bold().FontColor(Colors.White.Default));
                            row.Cell().Text("Description", t => t.Bold().FontColor(Colors.White.Default));
                            row.Cell().Text("Qty", t => t.Bold().FontColor(Colors.White.Default).AlignRight());
                            row.Cell().Text("Barcode", t => t.Bold().FontColor(Colors.White.Default));
                        });

                        foreach (var package in Packages)
                        {
                            table.Row(row =>
                            {
                                row.Cell().VerticalAlignMiddle().Text(package.Sku);
                                row.Cell().VerticalAlignMiddle().Text(package.Description);
                                row.Cell().VerticalAlignMiddle().Text(package.Qty.ToString(), t => t.AlignRight());
                                row.Cell().VerticalAlignMiddle().Barcode(package.Sku, bc => bc
                                    .Width(110)
                                    .Height(22)
                                    .ShowText(false)
                                    .AltText($"Barcode for {package.Sku}"));
                            });
                        }
                    });

                    col.Item().H2("Package Labels");
                    col.Item().Text("Standalone labels, sized and styled for printing on shipping boxes.");
                    col.Item().Row(row =>
                    {
                        row.Spacing(12);
                        row.RelativeItem().Column(label =>
                        {
                            label.Item().Text("Fragile - Handle With Care", t => t.Bold().AlignCenter());
                            label.Item().Barcode("FRAGILE-01", bc => bc
                                .Width(160)
                                .Height(45)
                                .BarColor(Colors.Red.L700)
                                .AlignCenter()
                                .Caption("Handling code FRAGILE-01"));
                        });

                        row.RelativeItem().Column(label =>
                        {
                            label.Item().Text("Return To Sender", t => t.Bold().AlignCenter());
                            label.Item().Barcode("RETURN-WH04", bc => bc
                                .Width(160)
                                .Height(45)
                                .BarColor(Colors.Grey.L800)
                                .AlignCenter()
                                .Caption("Warehouse WH-04"));
                        });
                    });

                    col.Item().H3("About These Barcodes");
                    col.Item().BulletList(list =>
                    {
                        list.Item("Encoded as Code 128 (Subset B), which supports the full printable ASCII range.");
                        list.Item("Rendered as vector shapes rather than raster images, so they stay crisp when zoomed, printed, or resized.");
                        list.Item("Configurable size, bar color, alignment, an optional human-readable text line, and captions.");
                        list.Item("Usable anywhere an image can go, including table cells, as shown in the package contents table above.");
                    });
                });
            });
        }).PublishDocx(documentPath);

        return documentPath;
    }
}
