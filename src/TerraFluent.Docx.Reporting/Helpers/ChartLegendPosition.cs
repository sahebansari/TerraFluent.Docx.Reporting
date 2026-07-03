namespace TerraFluent.Docx.Reporting;

/// <summary>
/// Where a chart's legend is placed, applied via <see cref="Infra.IChartDescriptor.Legend"/>.
/// </summary>
public enum ChartLegendPosition
{
    /// <summary>Legend to the right of the plot area (Word's default).</summary>
    Right,
    /// <summary>Legend to the left of the plot area.</summary>
    Left,
    /// <summary>Legend above the plot area.</summary>
    Top,
    /// <summary>Legend below the plot area.</summary>
    Bottom
}
