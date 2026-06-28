namespace TerraFluent.Docx.Reporting;

/// <summary>
/// The shared font, color, and table-layout defaults for a document. Construct one directly to
/// reuse across documents, or configure it inline via
/// <see cref="Infra.IDocumentContainer.Theme(Action{Infra.IDocumentThemeDescriptor})"/>.
/// </summary>
public sealed class DocumentTheme
{
    /// <summary>The default font family for document text.</summary>
    public string DefaultFontFamily { get; set; } = "Calibri";
    /// <summary>The default font size, in points, for document text.</summary>
    public float DefaultFontSize { get; set; } = 11;
    /// <summary>The default body text color.</summary>
    public string DefaultTextColor { get; set; } = Colors.Grey.L900;
    /// <summary>The color used for H1-H6 headings.</summary>
    public string HeadingColor { get; set; } = Colors.Grey.L900;
    /// <summary>The document's accent color, used as a default for heading and table header colors.</summary>
    public string AccentColor { get; set; } = Colors.Blue.L700;
    /// <summary>The color used for hyperlink runs that don't override their own font color.</summary>
    public string HyperlinkColor { get; set; } = Colors.Blue.L700;
    /// <summary>The default table border line width, in points.</summary>
    public float TableBorderWidth { get; set; }
    /// <summary>The default table border color.</summary>
    public string TableBorderColor { get; set; } = Colors.Grey.L300;
    /// <summary>The default table cell padding, in points, on the top and bottom sides.</summary>
    public float TableCellPaddingVertical { get; set; } = 3;
    /// <summary>The default table cell padding, in points, on the left and right sides.</summary>
    public float TableCellPaddingHorizontal { get; set; } = 5;
    /// <summary>The default background fill color for table header rows, or <see langword="null"/> for no fill.</summary>
    public string? TableHeaderBackgroundColor { get; set; }
    /// <summary>The default background fill color for alternating table body rows, or <see langword="null"/> for no banding.</summary>
    public string? TableAlternateRowBackgroundColor { get; set; }
    /// <summary>The default minimum height, in points, for table body rows.</summary>
    public float TableRowMinHeight { get; set; }
    /// <summary>The default minimum height, in points, for table header rows.</summary>
    public float TableHeaderRowMinHeight { get; set; }

    /// <summary>Creates a new theme with default values.</summary>
    public static DocumentTheme Default => new();

    /// <summary>Creates an independent copy of this theme.</summary>
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
