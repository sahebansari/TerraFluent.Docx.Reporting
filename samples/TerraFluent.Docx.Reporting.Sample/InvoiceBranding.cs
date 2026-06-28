using TerraFluent.Docx.Reporting;
using TerraFluent.Docx.Reporting.Infra;

internal static class InvoiceBranding
{
    private const string CompanyName = "Northwind Consulting";
    private const string LogoAltText = "Northwind Consulting company logo";

    public static void Header(IContainer header, string logoPath, string documentLabel, float documentLabelSize)
    {
        header.Row(row =>
        {
            row.Spacing(6);
            row.ConstantItem(34).Image(logoPath, img => img
                .Width(28)
                .AltText(LogoAltText));

            row.RelativeItem().Column(brand =>
            {
                brand.Item().Text(CompanyName, t => t.Bold().FontSize(13).FontColor(Colors.Blue.L700));
                brand.Item().Text("Document automation services", t => t.FontSize(8.5f).FontColor(Colors.Grey.L600));
            });

            row.AutoItem().Text(documentLabel, t => t.Bold().FontSize(documentLabelSize).FontColor(Colors.Grey.L800).AlignRight());
        });
    }

    public static void Footer(IContainer footer, string logoPath, bool includePageNumbers)
    {
        footer.Column(col =>
        {
            col.Item().Line();
            col.Item().Row(row =>
            {
                row.Spacing(6);
                row.ConstantItem(24).Image(logoPath, img => img
                    .Width(18)
                    .AltText(LogoAltText));

                row.RelativeItem().Text("Thank you for your business.", t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignCenter());

                if (includePageNumbers)
                {
                    row.AutoItem().Text(t =>
                    {
                        t.Span("Page ").FontSize(9).FontColor(Colors.Grey.L600);
                        t.CurrentPageNumber(s => s.FontSize(9).FontColor(Colors.Grey.L600));
                        t.Span(" of ").FontSize(9).FontColor(Colors.Grey.L600);
                        t.TotalPages(s => s.FontSize(9).FontColor(Colors.Grey.L600));
                        t.AlignRight();
                    });
                }
                else
                {
                    row.AutoItem().Text(CompanyName, t => t.FontSize(9).FontColor(Colors.Grey.L600).AlignRight());
                }
            });
        });
    }
}
