namespace TerraFluent.Docx.Reporting;

/// <summary>
/// A page width and height, in points. Use one of the predefined sizes or construct a custom size directly.
/// </summary>
public readonly struct PageSize
{
    /// <summary>The page width, in points.</summary>
    public float Width { get; }

    /// <summary>The page height, in points.</summary>
    public float Height { get; }

    /// <summary>Creates a custom page size.</summary>
    /// <param name="width">Page width, in points.</param>
    /// <param name="height">Page height, in points.</param>
    public PageSize(float width, float height)
    {
        Width = width;
        Height = height;
    }

    // ISO A-series (in points: 1pt = 1/72 inch)
    /// <summary>ISO A0 (841 x 1189 mm), portrait orientation.</summary>
    public static PageSize A0 => new(2383.94f, 3370.39f);
    /// <summary>ISO A1 (594 x 841 mm), portrait orientation.</summary>
    public static PageSize A1 => new(1683.78f, 2383.94f);
    /// <summary>ISO A2 (420 x 594 mm), portrait orientation.</summary>
    public static PageSize A2 => new(1190.55f, 1683.78f);
    /// <summary>ISO A3 (297 x 420 mm), portrait orientation.</summary>
    public static PageSize A3 => new(841.89f, 1190.55f);
    /// <summary>ISO A4 (210 x 297 mm), portrait orientation.</summary>
    public static PageSize A4 => new(595.28f, 841.89f);
    /// <summary>ISO A5 (148 x 210 mm), portrait orientation.</summary>
    public static PageSize A5 => new(419.53f, 595.28f);
    /// <summary>ISO A6 (105 x 148 mm), portrait orientation.</summary>
    public static PageSize A6 => new(297.64f, 419.53f);

    // North American
    /// <summary>US Letter (8.5 x 11 in), portrait orientation.</summary>
    public static PageSize Letter => new(612f, 792f);
    /// <summary>US Legal (8.5 x 14 in), portrait orientation.</summary>
    public static PageSize Legal  => new(612f, 1008f);
    /// <summary>US Ledger (17 x 11 in), landscape orientation.</summary>
    public static PageSize Ledger => new(1224f, 792f);
    /// <summary>US Tabloid (11 x 17 in), portrait orientation.</summary>
    public static PageSize Tabloid => new(792f, 1224f);

    /// <summary>Returns this size rotated to landscape, swapping width and height.</summary>
    public PageSize Landscape() => new(Height, Width);
}
