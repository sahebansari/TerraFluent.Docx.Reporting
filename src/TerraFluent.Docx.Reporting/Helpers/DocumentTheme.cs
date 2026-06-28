namespace TerraFluent.Docx.Reporting;

public sealed class DocumentTheme
{
    public string DefaultFontFamily { get; set; } = "Calibri";
    public float DefaultFontSize { get; set; } = 11;
    public string DefaultTextColor { get; set; } = Colors.Grey.L900;
    public string HeadingColor { get; set; } = Colors.Grey.L900;
    public string AccentColor { get; set; } = Colors.Blue.L700;
    public string HyperlinkColor { get; set; } = Colors.Blue.L700;
    public float TableBorderWidth { get; set; }
    public string TableBorderColor { get; set; } = Colors.Grey.L300;
    public float TableCellPaddingVertical { get; set; } = 3;
    public float TableCellPaddingHorizontal { get; set; } = 5;
    public string? TableHeaderBackgroundColor { get; set; }
    public string? TableAlternateRowBackgroundColor { get; set; }
    public float TableRowMinHeight { get; set; }
    public float TableHeaderRowMinHeight { get; set; }

    public static DocumentTheme Default => new();

    public DocumentTheme Clone() => new()
    {
        DefaultFontFamily = DefaultFontFamily,
        DefaultFontSize = DefaultFontSize,
        DefaultTextColor = DefaultTextColor,
        HeadingColor = HeadingColor,
        AccentColor = AccentColor,
        HyperlinkColor = HyperlinkColor,
        TableBorderWidth = TableBorderWidth,
        TableBorderColor = TableBorderColor,
        TableCellPaddingVertical = TableCellPaddingVertical,
        TableCellPaddingHorizontal = TableCellPaddingHorizontal,
        TableHeaderBackgroundColor = TableHeaderBackgroundColor,
        TableAlternateRowBackgroundColor = TableAlternateRowBackgroundColor,
        TableRowMinHeight = TableRowMinHeight,
        TableHeaderRowMinHeight = TableHeaderRowMinHeight
    };

    /// <summary>
    /// Copies all property values from <paramref name="other"/> into this instance.
    /// Used to update a shared theme instance in place so descriptors created before
    /// and after the update see the same values.
    /// </summary>
    public void CopyFrom(DocumentTheme other)
    {
        DefaultFontFamily = other.DefaultFontFamily;
        DefaultFontSize = other.DefaultFontSize;
        DefaultTextColor = other.DefaultTextColor;
        HeadingColor = other.HeadingColor;
        AccentColor = other.AccentColor;
        HyperlinkColor = other.HyperlinkColor;
        TableBorderWidth = other.TableBorderWidth;
        TableBorderColor = other.TableBorderColor;
        TableCellPaddingVertical = other.TableCellPaddingVertical;
        TableCellPaddingHorizontal = other.TableCellPaddingHorizontal;
        TableHeaderBackgroundColor = other.TableHeaderBackgroundColor;
        TableAlternateRowBackgroundColor = other.TableAlternateRowBackgroundColor;
        TableRowMinHeight = other.TableRowMinHeight;
        TableHeaderRowMinHeight = other.TableHeaderRowMinHeight;
    }
}
