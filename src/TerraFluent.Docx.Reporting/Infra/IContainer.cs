namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Adds content to a region of a document - the page body, a header or footer, a column, a row item,
/// or a table cell. Most fluent document building happens through this interface.
/// </summary>
public interface IContainer
{
    /// <summary>Adds a level 1 heading paragraph (Word style "Heading1").</summary>
    IContainer H1(string text);

    /// <summary>Adds a level 2 heading paragraph (Word style "Heading2").</summary>
    IContainer H2(string text);

    /// <summary>Adds a level 3 heading paragraph (Word style "Heading3").</summary>
    IContainer H3(string text);

    /// <summary>Adds a level 4 heading paragraph (Word style "Heading4").</summary>
    IContainer H4(string text);

    /// <summary>Adds a level 5 heading paragraph (Word style "Heading5").</summary>
    IContainer H5(string text);

    /// <summary>Adds a level 6 heading paragraph (Word style "Heading6").</summary>
    IContainer H6(string text);

    /// <summary>
    /// Adds a vertical stack of content. Use <see cref="IColumnDescriptor.Item"/> to add each block.
    /// </summary>
    /// <param name="configure">Configures the column's spacing and items. Cannot be null.</param>
    IContainer Column(Action<IColumnDescriptor> configure);

    /// <summary>
    /// Adds a horizontal layout row made up of side-by-side item containers.
    /// </summary>
    /// <param name="configure">Configures the row's spacing and items. Cannot be null.</param>
    IContainer Row(Action<IRowDescriptor> configure);

    /// <summary>Adds a paragraph with the given plain text and optional formatting.</summary>
    /// <param name="text">The paragraph text.</param>
    /// <param name="configure">Optional formatting for the paragraph, such as alignment or font.</param>
    IContainer Text(string text, Action<ITextDescriptor>? configure = null);

    /// <summary>
    /// Adds a paragraph built from multiple formatted runs via <see cref="ITextDescriptor.Span"/>.
    /// </summary>
    /// <param name="configure">Builds the paragraph's runs and formatting. Cannot be null.</param>
    IContainer Text(Action<ITextDescriptor> configure);

    /// <summary>Adds a paragraph consisting of a single hyperlink.</summary>
    /// <param name="text">The visible link text.</param>
    /// <param name="url">The link target URL.</param>
    /// <param name="configure">Optional formatting for the hyperlink run.</param>
    IContainer Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null);

    /// <summary>Adds an invisible bookmark anchor that can be targeted by <see cref="ITextDescriptor.CrossReference"/>.</summary>
    /// <param name="name">The bookmark name, unique within the document.</param>
    IContainer Bookmark(string name);

    /// <summary>
    /// Adds a visible paragraph that also carries a bookmark, so it can be targeted by a cross reference.
    /// <see cref="H1"/>-<see cref="H6"/> don't accept a bookmark name; to bookmark a heading, call this
    /// overload and apply a heading style with <c>configure: t =&gt; t.Style("Heading1")</c> (through
    /// <c>"Heading6"</c>) instead of using <see cref="H1"/>-<see cref="H6"/>.
    /// </summary>
    /// <param name="name">The bookmark name, unique within the document.</param>
    /// <param name="text">The paragraph text.</param>
    /// <param name="configure">Optional formatting for the paragraph.</param>
    IContainer Bookmark(string name, string text, Action<ITextDescriptor>? configure = null);

    /// <summary>Adds a Word table-of-contents field that Word populates from heading styles when the field is updated.</summary>
    /// <param name="title">The heading text shown above the table of contents.</param>
    /// <param name="minLevel">The lowest heading level (1-9) to include.</param>
    /// <param name="maxLevel">The highest heading level (1-9) to include.</param>
    IContainer TableOfContents(string title = "Contents", int minLevel = 1, int maxLevel = 3);

    /// <summary>Adds an unordered (bulleted) list.</summary>
    /// <param name="configure">Adds items and configures markers. Cannot be null.</param>
    IContainer BulletList(Action<IListDescriptor> configure);

    /// <summary>Adds an ordered (numbered) list.</summary>
    /// <param name="configure">Adds items and configures markers. Cannot be null.</param>
    IContainer NumberedList(Action<IListDescriptor> configure);

    /// <summary>Adds a table.</summary>
    /// <param name="configure">Defines columns, header row, body rows, and styling. Cannot be null.</param>
    IContainer Table(Action<ITableDescriptor> configure);

    /// <summary>Adds a chart rendered as a native Word chart part.</summary>
    /// <param name="configure">Sets the chart title and data series. Cannot be null.</param>
    IContainer Chart(Action<IChartDescriptor> configure);

    /// <summary>Adds an inline image loaded from a file path, resolved when the document is published.</summary>
    /// <param name="filePath">Path to the image file. Must exist when the document is published.</param>
    /// <param name="width">Optional width in points. The aspect ratio is preserved.</param>
    IContainer Image(string filePath, float? width = null);

    /// <summary>Adds an image loaded from a file path, with full sizing, wrapping, and styling options.</summary>
    /// <param name="filePath">Path to the image file. Must exist when the document is published.</param>
    /// <param name="configure">Configures size, alignment, wrapping, borders, and other image options. Cannot be null.</param>
    IContainer Image(string filePath, Action<IImageDescriptor> configure);

    /// <summary>Adds an image from an in-memory byte array.</summary>
    /// <param name="imageBytes">The encoded image bytes (e.g. PNG or JPEG). Cannot be null or empty.</param>
    /// <param name="fileName">A file name used to infer the image format, e.g. "chart.png".</param>
    /// <param name="configure">Optional configuration for size, alignment, wrapping, and styling.</param>
    IContainer Image(byte[] imageBytes, string fileName, Action<IImageDescriptor>? configure = null);

    /// <summary>Adds an image read from a stream. The stream is copied and is not closed by this method.</summary>
    /// <param name="imageStream">A readable stream positioned at the start of the encoded image.</param>
    /// <param name="fileName">A file name used to infer the image format, e.g. "chart.png".</param>
    /// <param name="configure">Optional configuration for size, alignment, wrapping, and styling.</param>
    IContainer Image(Stream imageStream, string fileName, Action<IImageDescriptor>? configure = null);

    /// <summary>Adds a horizontal rule (a paragraph with a bottom border).</summary>
    IContainer Line();

    /// <summary>Adds an explicit page break.</summary>
    IContainer PageBreak();

    /// <summary>Composes a reusable <see cref="IComponent"/> into this container.</summary>
    /// <param name="component">The component to render. Cannot be null.</param>
    IContainer Component(IComponent component);
}
