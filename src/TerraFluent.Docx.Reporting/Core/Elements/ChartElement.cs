namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class ChartElement : IElement
{
    public string Title { get; set; } = "Chart";
    public List<ChartSeriesElement> Series { get; } = [];
    public bool ShowLegend { get; set; } = true;
    public string LegendPosition { get; set; } = "r";  // OOXML ST_LegendPos: r, l, t, b
    public string? CategoryAxisTitle { get; set; }
    public string? ValueAxisTitle { get; set; }
    public bool ShowDataLabels { get; set; }
    public string BarGrouping { get; set; } = "clustered";  // OOXML ST_BarGrouping: clustered, stacked, percentStacked
    public float? Width { get; set; }   // points; default 432 (6 in)
    public float? Height { get; set; }  // points; default 252 (3.5 in)
    public string? Alignment { get; set; }
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
