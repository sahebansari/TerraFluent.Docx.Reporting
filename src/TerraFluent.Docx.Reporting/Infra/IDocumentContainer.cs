namespace TerraFluent.Docx.Reporting.Infra;

public interface IDocumentContainer
{
    IDocumentContainer Theme(DocumentTheme theme);
    IDocumentContainer Theme(Action<IDocumentThemeDescriptor> configure);
    IDocumentContainer ParagraphStyle(string name, Action<ITextDescriptor> configure);
    IDocumentContainer TableStyle(string name, Action<ITableDescriptor> configure);
    IDocumentContainer Page(Action<IPageDescriptor> configure);
    IDocumentContainer MetadataTitle(string title);
    IDocumentContainer MetadataAuthor(string author);
    IDocumentContainer MetadataSubject(string subject);
    IDocumentContainer MetadataKeywords(string keywords);
    IDocumentContainer MetadataCreator(string creator);
    IDocumentContainer Compose(IDocument document);
}
