namespace TerraFluent.Docx.Reporting;

/// <summary>
/// Common six-digit hex color constants, grouped by hue with Material-Design-style shade numbers
/// (lower numbers are lighter). Any API that accepts a color also accepts any six-digit hex string
/// or three-digit CSS shorthand, with or without a leading "#" - these constants are just convenient,
/// named values.
/// </summary>
public static class Colors
{
    /// <summary>Solid black.</summary>
    public static class Black
    {
        /// <summary>Solid black (000000).</summary>
        public const string Default = "000000";
    }

    /// <summary>Solid white.</summary>
    public static class White
    {
        /// <summary>Solid white (FFFFFF).</summary>
        public const string Default = "FFFFFF";
    }

    /// <summary>Greyscale shades from lightest (L100) to darkest (L900).</summary>
    public static class Grey
    {
        /// <summary>Lightest grey shade.</summary>
        public const string L100 = "F5F5F5";
        /// <summary>Grey shade, lighter than <see cref="L300"/>.</summary>
        public const string L200 = "EEEEEE";
        /// <summary>Grey shade, lighter than <see cref="L400"/>.</summary>
        public const string L300 = "E0E0E0";
        /// <summary>Grey shade, lighter than <see cref="L500"/>.</summary>
        public const string L400 = "BDBDBD";
        /// <summary>Mid grey shade.</summary>
        public const string L500 = "9E9E9E";
        /// <summary>Grey shade, darker than <see cref="L500"/>.</summary>
        public const string L600 = "757575";
        /// <summary>Grey shade, darker than <see cref="L600"/>.</summary>
        public const string L700 = "616161";
        /// <summary>Grey shade, darker than <see cref="L700"/>.</summary>
        public const string L800 = "424242";
        /// <summary>Darkest grey shade.</summary>
        public const string L900 = "212121";
    }

    /// <summary>Blue shades from lightest (L100) to darkest (L900).</summary>
    public static class Blue
    {
        /// <summary>Lightest blue shade.</summary>
        public const string L100 = "BBDEFB";
        /// <summary>Blue shade, lighter than <see cref="L300"/>.</summary>
        public const string L200 = "90CAF9";
        /// <summary>Blue shade, lighter than <see cref="L400"/>.</summary>
        public const string L300 = "64B5F6";
        /// <summary>Blue shade, lighter than <see cref="L500"/>.</summary>
        public const string L400 = "42A5F5";
        /// <summary>Mid blue shade.</summary>
        public const string L500 = "2196F3";
        /// <summary>Blue shade, darker than <see cref="L500"/>.</summary>
        public const string L600 = "1E88E5";
        /// <summary>Blue shade, darker than <see cref="L600"/>.</summary>
        public const string L700 = "1976D2";
        /// <summary>Blue shade, darker than <see cref="L700"/>.</summary>
        public const string L800 = "1565C0";
        /// <summary>Darkest blue shade.</summary>
        public const string L900 = "0D47A1";
    }

    /// <summary>Red shades from lightest (L100) to darkest (L900).</summary>
    public static class Red
    {
        /// <summary>Lightest red shade.</summary>
        public const string L100 = "FFCDD2";
        /// <summary>Mid red shade.</summary>
        public const string L500 = "F44336";
        /// <summary>Red shade, darker than <see cref="L500"/>.</summary>
        public const string L700 = "D32F2F";
        /// <summary>Darkest red shade.</summary>
        public const string L900 = "B71C1C";
    }

    /// <summary>Green shades from lightest (L100) to darkest (L900).</summary>
    public static class Green
    {
        /// <summary>Lightest green shade.</summary>
        public const string L100 = "C8E6C9";
        /// <summary>Mid green shade.</summary>
        public const string L500 = "4CAF50";
        /// <summary>Green shade, darker than <see cref="L500"/>.</summary>
        public const string L700 = "388E3C";
        /// <summary>Darkest green shade.</summary>
        public const string L900 = "1B5E20";
    }

    /// <summary>Orange shades from lightest (L100) to darkest (L900).</summary>
    public static class Orange
    {
        /// <summary>Lightest orange shade.</summary>
        public const string L100 = "FFE0B2";
        /// <summary>Mid orange shade.</summary>
        public const string L500 = "FF9800";
        /// <summary>Orange shade, darker than <see cref="L500"/>.</summary>
        public const string L700 = "F57C00";
        /// <summary>Darkest orange shade.</summary>
        public const string L900 = "E65100";
    }
}
