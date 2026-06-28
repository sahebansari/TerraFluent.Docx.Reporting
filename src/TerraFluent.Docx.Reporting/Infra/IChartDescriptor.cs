namespace TerraFluent.Docx.Reporting.Infra;

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
}

