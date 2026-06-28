namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class ChartElement : IElement
{
    public string Title { get; set; } = "Chart";
    public List<ChartSeriesElement> Series { get; } = [];
}

internal sealed class ChartSeriesElement
{
    public string Name { get; set; } = string.Empty;
    public string Kind { get; set; } = "bar";
    public List<ChartPointElement> Points { get; } = [];
    public string Color { get; set; } = Colors.Blue.L700;
}

internal sealed class ChartPointElement
{
    public string Label { get; set; } = string.Empty;
    public double Value { get; set; }
}
