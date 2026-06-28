namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Fluent descriptor for adding data points to a chart series.
/// </summary>
public interface ISeriesDescriptor
{
    /// <summary>
    /// Adds a bar chart data point to this series.
    /// </summary>
    ISeriesDescriptor Bar(string label, double value);

    /// <summary>
    /// Adds a line chart data point to this series.
    /// </summary>
    ISeriesDescriptor Line(string label, double value);

    /// <summary>
    /// Adds a pie chart data point to this series.
    /// </summary>
    ISeriesDescriptor Pie(string label, double value);

    /// <summary>
    /// Adds a doughnut chart data point to this series.
    /// </summary>
    ISeriesDescriptor Doughnut(string label, double value);

    /// <summary>
    /// Sets the color for this series (hex color code, e.g., "FF0000" for red).
    /// </summary>
    ISeriesDescriptor Color(string hexColor);
}
