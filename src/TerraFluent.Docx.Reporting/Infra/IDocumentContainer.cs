namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// The root builder for a document, received by the configuration callback passed to
/// <see cref="Document.Create"/>. Sets document-wide theme and metadata, registers reusable
/// styles, and adds pages.
/// </summary>
public interface IDocumentContainer
{
    /// <summary>Applies an existing, possibly shared, theme object to the document.</summary>
    /// <param name="theme">The theme to apply.</param>
    IDocumentContainer Theme(DocumentTheme theme);

    /// <summary>Configures the document's theme inline.</summary>
    /// <param name="configure">Sets fonts, colors, and table defaults. Cannot be null.</param>
    IDocumentContainer Theme(Action<IDocumentThemeDescriptor> configure);

    /// <summary>Registers a reusable named paragraph style that can be applied via <see cref="ITextDescriptor.Style"/>.</summary>
    /// <param name="name">The style name used to reference it later.</param>
    /// <param name="configure">Configures the style's formatting. Cannot be null.</param>
    IDocumentContainer ParagraphStyle(string name, Action<ITextDescriptor> configure);

    /// <summary>Registers a reusable named table style that can be applied via <see cref="ITableDescriptor.Style"/>.</summary>
    /// <param name="name">The style name used to reference it later.</param>
    /// <param name="configure">Configures the style's formatting. Cannot be null.</param>
    IDocumentContainer TableStyle(string name, Action<ITableDescriptor> configure);

    /// <summary>Adds a page (Word section) to the document.</summary>
    /// <param name="configure">Configures the page's size, margins, headers, footers, and content. Cannot be null.</param>
    IDocumentContainer Page(Action<IPageDescriptor> configure);

    /// <summary>Sets the document's title metadata property.</summary>
    IDocumentContainer MetadataTitle(string title);

    /// <summary>Sets the document's author metadata property.</summary>
    IDocumentContainer MetadataAuthor(string author);

    /// <summary>Sets the document's subject metadata property.</summary>
    IDocumentContainer MetadataSubject(string subject);

    /// <summary>Sets the document's keywords metadata property.</summary>
    IDocumentContainer MetadataKeywords(string keywords);

    /// <summary>Sets the document's creator (application) metadata property.</summary>
    IDocumentContainer MetadataCreator(string creator);

    /// <summary>Composes a reusable <see cref="IDocument"/> module into this document.</summary>
    /// <param name="document">The document module to add. Cannot be null.</param>
    IDocumentContainer Compose(IDocument document);
}
