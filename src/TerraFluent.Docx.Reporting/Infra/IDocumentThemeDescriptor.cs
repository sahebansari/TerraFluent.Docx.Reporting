namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Configures a document's shared defaults: fonts, text and table colors, and table layout
/// defaults. Reached via <see cref="IDocumentContainer.Theme(Action{IDocumentThemeDescriptor})"/>.
/// Values set here are read eagerly by descriptors created afterward; calling <c>Theme</c> again
/// later updates the shared theme instance, so content added even earlier in document order also
/// picks up the new values for properties it reads lazily (e.g. hyperlink color).
/// </summary>
public interface IDocumentThemeDescriptor
{
    /// <summary>Sets both the default font family and size in one call.</summary>
    /// <param name="family">The default font family name.</param>
    /// <param name="size">The default font size, in points.</param>
    IDocumentThemeDescriptor DefaultFont(string family, float size = 11);

    /// <summary>Sets only the default font family, leaving the default size unchanged.</summary>
    IDocumentThemeDescriptor DefaultFontFamily(string family);

    /// <summary>Sets only the default font size, in points, leaving the default family unchanged.</summary>
    IDocumentThemeDescriptor DefaultFontSize(float size);

    /// <summary>Sets the default body text color.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IDocumentThemeDescriptor DefaultTextColor(string hexColor);

    /// <summary>Sets the color used for H1-H6 headings.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IDocumentThemeDescriptor HeadingColor(string hexColor);

    /// <summary>
    /// Sets the document's accent color. This also sets <see cref="HeadingColor"/> to the same
    /// color, and sets the default table header background to it if no table header background
    /// has been set yet.
    /// </summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IDocumentThemeDescriptor AccentColor(string hexColor);

    /// <summary>Sets the color used for hyperlink runs that don't override their own font color.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IDocumentThemeDescriptor HyperlinkColor(string hexColor);

    /// <summary>Sets the default background fill color for table header rows.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IDocumentThemeDescriptor TableHeaderBackground(string hexColor);

    /// <summary>Sets the default background fill color applied to alternating table body rows.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IDocumentThemeDescriptor TableAlternateRowBackground(string hexColor);

    /// <summary>Sets the default table border style.</summary>
    /// <param name="width">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IDocumentThemeDescriptor TableBorder(float width, string hexColor);

    /// <summary>Sets uniform default table cell padding, in points, on all sides.</summary>
    IDocumentThemeDescriptor TableCellPadding(float points);

    /// <summary>Sets default table cell padding, in points, separately for top/bottom and left/right.</summary>
    IDocumentThemeDescriptor TableCellPadding(float verticalPoints, float horizontalPoints);

    /// <summary>Sets the default minimum height, in points, for table body rows.</summary>
    IDocumentThemeDescriptor TableRowMinHeight(float points);

    /// <summary>Sets the default minimum height, in points, for table header rows.</summary>
    IDocumentThemeDescriptor TableHeaderRowMinHeight(float points);
}
