namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Configures a paragraph's text and formatting. Methods that affect a single run of text (e.g.
/// <see cref="Bold"/>, <see cref="FontColor(string)"/>) apply to the most recently added run -
/// the whole paragraph until <see cref="Span"/> starts a new run. Methods that affect the whole
/// paragraph (e.g. <see cref="AlignLeft"/>, <see cref="SpacingBefore"/>) always apply to the
/// paragraph, even when called on the descriptor returned by <see cref="Span"/>.
/// </summary>
public interface ITextDescriptor
{
    /// <summary>Sets or clears bold for the current run.</summary>
    ITextDescriptor Bold(bool value = true);

    /// <summary>Sets or clears italic for the current run.</summary>
    ITextDescriptor Italic(bool value = true);

    /// <summary>Sets or clears underline for the current run.</summary>
    ITextDescriptor Underline(bool value = true);

    /// <summary>Sets or clears strikethrough for the current run.</summary>
    ITextDescriptor Strikethrough(bool value = true);

    /// <summary>Sets or clears superscript for the current run. Mutually exclusive with <see cref="Subscript"/>.</summary>
    ITextDescriptor Superscript(bool value = true);

    /// <summary>Sets or clears subscript for the current run. Mutually exclusive with <see cref="Superscript"/>.</summary>
    ITextDescriptor Subscript(bool value = true);

    /// <summary>Sets the font size, in points, for the current run.</summary>
    ITextDescriptor FontSize(float size);

    /// <summary>Sets the font color for the current run.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITextDescriptor FontColor(string hexColor);

    /// <summary>
    /// Sets the font family for the current run. This only sets the Latin/East Asian font
    /// (<c>w:ascii</c>/<c>w:hAnsi</c>); it does not set a complex-script font (<c>w:cs</c>), so
    /// right-to-left or complex-script text in this run still renders with the document's default font.
    /// </summary>
    ITextDescriptor FontFamily(string name);

    /// <summary>Applies a text highlight color to the current run.</summary>
    ITextDescriptor Highlight(HighlightColor color = HighlightColor.Yellow);

    /// <summary>Applies a registered paragraph style by name. The style must have been registered via
    /// <see cref="IDocumentContainer.ParagraphStyle"/>.</summary>
    ITextDescriptor Style(string name);

    /// <summary>Left-aligns the paragraph.</summary>
    ITextDescriptor AlignLeft();

    /// <summary>Center-aligns the paragraph.</summary>
    ITextDescriptor AlignCenter();

    /// <summary>Right-aligns the paragraph.</summary>
    ITextDescriptor AlignRight();

    /// <summary>Justifies the paragraph (stretches lines to fill the line width).</summary>
    ITextDescriptor Justify();

    /// <summary>Sets the line height as a multiplier of single spacing (e.g. 1.5 for 1.5 line spacing).</summary>
    ITextDescriptor LineHeight(float multiplier);

    /// <summary>Sets the spacing, in points, before the paragraph.</summary>
    ITextDescriptor SpacingBefore(float points);

    /// <summary>Sets the spacing, in points, after the paragraph.</summary>
    ITextDescriptor SpacingAfter(float points);

    /// <summary>Sets the left indent, in points, for the whole paragraph.</summary>
    ITextDescriptor LeftIndent(float points);

    /// <summary>Sets the right indent, in points, for the whole paragraph.</summary>
    ITextDescriptor RightIndent(float points);

    /// <summary>Sets an additional indent, in points, applied only to the paragraph's first line.</summary>
    ITextDescriptor FirstLineIndent(float points);

    /// <summary>Sets an additional indent, in points, applied to every line except the first (a hanging indent).</summary>
    ITextDescriptor HangingIndent(float points);

    /// <summary>Keeps this paragraph on the same page as the paragraph that follows it.</summary>
    ITextDescriptor KeepWithNext(bool value = true);

    /// <summary>Prevents this paragraph's lines from splitting across a page break.</summary>
    ITextDescriptor KeepLinesTogether(bool value = true);

    /// <summary>Forces this paragraph to start on a new page.</summary>
    ITextDescriptor PageBreakBefore(bool value = true);

    /// <summary>Sets the paragraph's background shading color.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITextDescriptor Shading(string hexColor);

    /// <summary>Sets a border on all four sides of the paragraph.</summary>
    /// <param name="widthPoints">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    /// <param name="spacePoints">Spacing, in points, between the border and the paragraph text.</param>
    ITextDescriptor Border(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);

    /// <summary>Sets a border on the top of the paragraph.</summary>
    /// <param name="widthPoints">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    /// <param name="spacePoints">Spacing, in points, between the border and the paragraph text.</param>
    ITextDescriptor BorderTop(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);

    /// <summary>Sets a border on the right of the paragraph.</summary>
    /// <param name="widthPoints">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    /// <param name="spacePoints">Spacing, in points, between the border and the paragraph text.</param>
    ITextDescriptor BorderRight(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);

    /// <summary>Sets a border on the bottom of the paragraph.</summary>
    /// <param name="widthPoints">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    /// <param name="spacePoints">Spacing, in points, between the border and the paragraph text.</param>
    ITextDescriptor BorderBottom(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);

    /// <summary>Sets a border on the left of the paragraph.</summary>
    /// <param name="widthPoints">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    /// <param name="spacePoints">Spacing, in points, between the border and the paragraph text.</param>
    ITextDescriptor BorderLeft(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);

    /// <summary>Adds a custom tab stop to the paragraph.</summary>
    /// <param name="positionPoints">The tab stop position, in points, from the left margin.</param>
    /// <param name="alignment">How text is aligned at the tab stop.</param>
    ITextDescriptor TabStop(float positionPoints, TabStopAlignment alignment = TabStopAlignment.Left);

    /// <summary>
    /// Starts a new formatted run within the paragraph and returns a descriptor scoped to it.
    /// Run-level formatting methods (e.g. <see cref="Bold"/>) called on the result apply only to
    /// this run; paragraph-level methods (e.g. <see cref="AlignLeft"/>) still apply to the whole paragraph.
    /// </summary>
    /// <param name="text">The run's text.</param>
    /// <param name="configure">Optional formatting for the run.</param>
    ITextDescriptor Span(string text, Action<ITextDescriptor>? configure = null);

    /// <summary>Inserts a tab character as its own run.</summary>
    ITextDescriptor Tab();

    /// <summary>Inserts a hyperlink run within the paragraph.</summary>
    /// <param name="text">The visible link text.</param>
    /// <param name="url">The link target URL.</param>
    /// <param name="configure">Optional formatting for the hyperlink run.</param>
    ITextDescriptor Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null);

    /// <summary>
    /// Inserts a field that displays the visible text of the paragraph or heading bookmarked with
    /// <paramref name="bookmarkName"/>, updating automatically if the target text changes.
    /// </summary>
    /// <param name="bookmarkName">The name of the bookmark to reference, as passed to <see cref="IContainer.Bookmark(string)"/>.</param>
    /// <param name="fallbackText">Text shown before the field is updated by Word. Defaults to the bookmark name.</param>
    /// <param name="configure">Optional formatting for the field run.</param>
    ITextDescriptor CrossReference(string bookmarkName, string? fallbackText = null, Action<ITextDescriptor>? configure = null);

    /// <summary>Inserts a footnote reference, with the given text as the note body.</summary>
    /// <param name="text">The footnote body text.</param>
    /// <param name="configure">Optional formatting for the footnote body's paragraph.</param>
    /// <exception cref="InvalidOperationException">Thrown if called from within another footnote or endnote body.</exception>
    ITextDescriptor Footnote(string text, Action<ITextDescriptor>? configure = null);

    /// <summary>Inserts an endnote reference, with the given text as the note body.</summary>
    /// <param name="text">The endnote body text.</param>
    /// <param name="configure">Optional formatting for the endnote body's paragraph.</param>
    /// <exception cref="InvalidOperationException">Thrown if called from within another footnote or endnote body.</exception>
    ITextDescriptor Endnote(string text, Action<ITextDescriptor>? configure = null);

    /// <summary>Inserts a field that displays the current page number, updated by Word when the field is refreshed.</summary>
    /// <param name="configure">Optional formatting for the field run.</param>
    ITextDescriptor CurrentPageNumber(Action<ITextDescriptor>? configure = null);

    /// <summary>Inserts a field that displays the total page count, updated by Word when the field is refreshed.</summary>
    /// <param name="configure">Optional formatting for the field run.</param>
    ITextDescriptor TotalPages(Action<ITextDescriptor>? configure = null);
}
