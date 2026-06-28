namespace TerraFluent.Docx.Reporting;

public readonly struct PageSize
{
    public float Width { get; }
    public float Height { get; }

    public PageSize(float width, float height)
    {
        Width = width;
        Height = height;
    }

    // ISO A-series (in points: 1pt = 1/72 inch)
    public static PageSize A0 => new(2383.94f, 3370.39f);
    public static PageSize A1 => new(1683.78f, 2383.94f);
    public static PageSize A2 => new(1190.55f, 1683.78f);
    public static PageSize A3 => new(841.89f, 1190.55f);
    public static PageSize A4 => new(595.28f, 841.89f);
    public static PageSize A5 => new(419.53f, 595.28f);
    public static PageSize A6 => new(297.64f, 419.53f);

    // North American
    public static PageSize Letter => new(612f, 792f);
    public static PageSize Legal  => new(612f, 1008f);
    public static PageSize Ledger => new(1224f, 792f);
    public static PageSize Tabloid => new(792f, 1224f);

    public PageSize Landscape() => new(Height, Width);
}
