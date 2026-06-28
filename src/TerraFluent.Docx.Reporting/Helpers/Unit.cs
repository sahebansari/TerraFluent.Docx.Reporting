namespace TerraFluent.Docx.Reporting;

/// <summary>
/// Converts common units to points, the layout unit used by TerraFluent.Docx.Reporting.
/// </summary>
public static class Unit
{
    /// <summary>
    /// Converts centimetres to points.
    /// </summary>
    public static float Centimetre(float value) => value * 28.3465f;

    /// <summary>
    /// Converts millimetres to points.
    /// </summary>
    public static float Millimetre(float value) => value * 2.83465f;

    /// <summary>
    /// Converts inches to points.
    /// </summary>
    public static float Inch(float value) => value * 72f;

    /// <summary>
    /// Returns the supplied point value.
    /// </summary>
    public static float Point(float value) => value;
}
