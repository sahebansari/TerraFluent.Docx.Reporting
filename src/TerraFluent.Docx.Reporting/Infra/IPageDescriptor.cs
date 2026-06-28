namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Configures a single page (Word section): size, margins, numbering, background, watermark,
/// columns, and the header/footer/content containers. Reached via <see cref="IDocumentContainer.Page"/>.
/// </summary>
public interface IPageDescriptor
{
    /// <summary>Sets the page size using a predefined or custom <see cref="PageSize"/>.</summary>
    IPageDescriptor Size(PageSize size);

    /// <summary>Sets the page width and height, in points.</summary>
    IPageDescriptor Size(float widthPoints, float heightPoints);

    /// <summary>Swaps width and height so the wider dimension becomes the page width.</summary>
    IPageDescriptor Landscape();

    /// <summary>Swaps width and height so the narrower dimension becomes the page width.</summary>
    IPageDescriptor Portrait();

    /// <summary>Sets a uniform margin, in points, on all four sides.</summary>
    IPageDescriptor Margin(float allPoints);

    /// <summary>Sets the top/bottom margin and the left/right margin, in points, independently.</summary>
    IPageDescriptor Margin(float verticalPoints, float horizontalPoints);

    /// <summary>Sets each margin, in points, independently.</summary>
    IPageDescriptor Margin(float topPoints, float rightPoints, float bottomPoints, float leftPoints);

    /// <summary>Sets only the top margin, in points.</summary>
    IPageDescriptor MarginTop(float points);

    /// <summary>Sets only the right margin, in points.</summary>
    IPageDescriptor MarginRight(float points);

    /// <summary>Sets only the bottom margin, in points.</summary>
    IPageDescriptor MarginBottom(float points);

    /// <summary>Sets only the left margin, in points.</summary>
    IPageDescriptor MarginLeft(float points);

    /// <summary>Sets the default paragraph formatting (font, size, color) inherited by content in this section
    /// that doesn't override it. Overrides the document theme's defaults for this section only.</summary>
    /// <param name="configure">Configures the default style. Cannot be null.</param>
    IPageDescriptor DefaultTextStyle(Action<ITextDescriptor> configure);

    /// <summary>Sets the starting page number for this section.</summary>
    /// <param name="value">The first page number, must be 1 or greater.</param>
    IPageDescriptor PageNumberStart(int value);

    /// <summary>
    /// Sets the page number format for this section to an OOXML ST_NumberFormat value
    /// (e.g., "lowerRoman" for i, ii, iii). The value is written as-is to the
    /// <c>w:fmt</c> attribute, so it must be one of the values defined by that
    /// enumeration or Word will treat the document as invalid.
    /// Common formats: "decimal" (default), "lowerRoman", "upperRoman", "lowerLetter", "upperLetter", "none".
    /// </summary>
    IPageDescriptor PageNumberFormat(string format);

    /// <summary>
    /// Sets the document's page background color (Word's "Page Color"). This maps to the
    /// single, document-wide &lt;w:background&gt; element in OOXML - there is no per-section
    /// or per-page equivalent, so the color cannot vary across pages within one .docx.
    /// If multiple pages call this, only the first non-empty value (in page order) is used
    /// and later calls are ignored.
    /// </summary>
    IPageDescriptor Background(string hexColor);

    /// <summary>Adds diagonal watermark text repeated across the page background.</summary>
    /// <param name="text">The watermark text.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    /// <param name="fontSize">The watermark font size, in points.</param>
    IPageDescriptor Watermark(string text, string hexColor = "D9D9D9", float fontSize = 54);

    /// <summary>Lays out the section's content body in newspaper-style columns.</summary>
    /// <param name="count">The number of columns, from 1 to 45.</param>
    /// <param name="spacingPoints">The spacing, in points, between columns.</param>
    /// <param name="separatorLine">Whether to draw a vertical line between columns.</param>
    IPageDescriptor Columns(int count, float spacingPoints = 36, bool separatorLine = false);

    /// <summary>Resets the section to a single column with the default spacing and no separator line.</summary>
    IPageDescriptor SingleColumn();

    /// <summary>Returns the section's default (odd-page) header container. Same as <see cref="OddPageHeader"/>.</summary>
    IContainer Header();

    /// <summary>Returns the section's odd-page header container. Same as <see cref="Header"/>.</summary>
    IContainer OddPageHeader();

    /// <summary>
    /// Returns the section's first-page header container, shown only on the section's first page.
    /// Adding content here enables Word's "different first page" header/footer behavior for the section.
    /// </summary>
    IContainer FirstPageHeader();

    /// <summary>
    /// Returns the section's even-page header container. Adding content here enables Word's
    /// "different odd and even pages" header/footer behavior for the whole document.
    /// </summary>
    IContainer EvenPageHeader();

    /// <summary>Returns the section's main body content container.</summary>
    IContainer Content();

    /// <summary>Returns the section's default (odd-page) footer container. Same as <see cref="OddPageFooter"/>.</summary>
    IContainer Footer();

    /// <summary>Returns the section's odd-page footer container. Same as <see cref="Footer"/>.</summary>
    IContainer OddPageFooter();

    /// <summary>
    /// Returns the section's first-page footer container, shown only on the section's first page.
    /// Adding content here enables Word's "different first page" header/footer behavior for the section.
    /// </summary>
    IContainer FirstPageFooter();

    /// <summary>
    /// Returns the section's even-page footer container. Adding content here enables Word's
    /// "different odd and even pages" header/footer behavior for the whole document.
    /// </summary>
    IContainer EvenPageFooter();
}
