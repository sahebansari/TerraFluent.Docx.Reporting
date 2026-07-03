namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Configures a chart added via <see cref="IContainer.Chart"/>: its title and data series. Charts
/// render as native Word chart parts, editable in Word like any chart it creates itself.
/// </summary>
public interface IChartDescriptor
{
    /// <summary>
    /// Sets the chart title.
    /// </summary>
    IChartDescriptor Title(string title);

    /// <summary>
    /// Adds a data series to the chart. For single-series charts, omit the name.
    /// All series in a chart must use the same chart type (all bar, all line, all pie, or all doughnut).
    /// </summary>
    /// <param name="configure">Configure the series by adding data points and optionally setting a color</param>
    IChartDescriptor Series(Action<ISeriesDescriptor> configure);

    /// <summary>
    /// Adds a named data series to the chart (for multi-series charts, where the name appears in the legend).
    /// All series in a chart must use the same chart type.
    /// </summary>
    /// <param name="name">Series name (displayed in legend)</param>
    /// <param name="configure">Configure the series by adding data points and optionally setting a color</param>
    IChartDescriptor Series(string name, Action<ISeriesDescriptor> configure);

    /// <summary>Sets the rendered width of the chart frame, in points. The default is 432 points (6 inches).</summary>
    IChartDescriptor Width(float points);

    /// <summary>Sets the rendered height of the chart frame, in points. The default is 252 points (3.5 inches).</summary>
    IChartDescriptor Height(float points);

    /// <summary>Left-aligns the chart's paragraph (the default).</summary>
    IChartDescriptor AlignLeft();

    /// <summary>Center-aligns the chart's paragraph.</summary>
    IChartDescriptor AlignCenter();

    /// <summary>Right-aligns the chart's paragraph.</summary>
    IChartDescriptor AlignRight();

    /// <summary>Places the legend on the given side of the plot area. The default is <see cref="ChartLegendPosition.Right"/>.</summary>
    IChartDescriptor Legend(ChartLegendPosition position);

    /// <summary>Hides the legend.</summary>
    IChartDescriptor HideLegend();

    /// <summary>Sets a title on the category (horizontal) axis. Ignored by pie and doughnut charts, which have no axes.</summary>
    IChartDescriptor CategoryAxisTitle(string title);

    /// <summary>Sets a title on the value (vertical) axis. Ignored by pie and doughnut charts, which have no axes.</summary>
    IChartDescriptor ValueAxisTitle(string title);

    /// <summary>Shows each data point's value as a label on the chart.</summary>
    IChartDescriptor DataLabels(bool show = true);

    /// <summary>
    /// Stacks the bars of a multi-series bar chart on top of each other instead of clustering them
    /// side by side. Only valid for bar charts.
    /// </summary>
    IChartDescriptor Stacked();

    /// <summary>
    /// Stacks the bars of a multi-series bar chart normalized to 100%, so each bar shows its series'
    /// share of the category total. Only valid for bar charts.
    /// </summary>
    IChartDescriptor PercentStacked();
}

