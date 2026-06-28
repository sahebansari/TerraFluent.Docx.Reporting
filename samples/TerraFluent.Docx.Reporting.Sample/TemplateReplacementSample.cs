using TerraFluent.Docx.Reporting;

internal static class TemplateReplacementSample
{
    public static string Generate(string outputDirectory)
    {
        var templatePath = Path.Combine(outputDirectory, "sample-template-source.docx");
        var outputPath = Path.Combine(outputDirectory, "sample-template-output.docx");

        Document.Create(container =>
        {
            container
                .MetadataTitle("Template Source")
                .MetadataAuthor("TerraFluent.Docx.Reporting Sample");

            container.Page(page =>
            {
                page.Size(PageSize.A4);
                page.Margin(Unit.Centimetre(2.54f));
                page.Content().H1("Client Brief");
                page.Content().Text("Prepared for {{ClientName}}");
                page.Content().Text("Project: {{ProjectName}}");
                page.Content().Text("Prepared by {{PreparedBy}}");
            });
        }).PublishDocx(templatePath);

        DocxTemplate.Open(templatePath)
            .Replace("ClientName", "Acme Industries")
            .Replace("ProjectName", "Document Automation Rollout")
            .Replace("PreparedBy", "Northwind Consulting")
            .SaveAs(outputPath);

        return outputPath;
    }
}
