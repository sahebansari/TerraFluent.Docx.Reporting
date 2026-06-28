var outputDirectory = SampleOutput.EnsureDesktopSampleDocs();
var imagePath = Path.Combine(outputDirectory, "sample-image.png");
var logoPath = Path.Combine(AppContext.BaseDirectory, "text-logo.png");

SampleImage.CreateGradientPng(imagePath, 64, 64);

var documents = new[]
{
    FeatureShowcaseSample.Generate(outputDirectory, imagePath),
    InvoiceSample.Generate(outputDirectory, logoPath),
    LongInvoiceSample.Generate(outputDirectory, logoPath),
    AnnualReportSample.Generate(outputDirectory, logoPath, imagePath),
    LayoutFeaturesSample.Generate(outputDirectory),
    TemplateReplacementSample.Generate(outputDirectory),
    ApiReferenceSample.Generate(outputDirectory, imagePath)
};

foreach (var document in documents)
    Console.WriteLine($"Created {document}");
