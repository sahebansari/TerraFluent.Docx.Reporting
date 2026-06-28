namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// A reusable, composable document module. Implement this to share theme, metadata, and page
/// definitions across documents, then add it to a builder via <see cref="IDocumentContainer.Compose"/>.
/// </summary>
public interface IDocument
{
    /// <summary>Adds this document's theme, metadata, and pages to the given document container.</summary>
    /// <param name="container">The document container to add content to.</param>
    void Compose(IDocumentContainer container);
}
