using TerraFluent.Docx.Reporting.Core;
using TerraFluent.Docx.Reporting.Infra;
using TerraFluent.Docx.Reporting.Ooxml;

namespace TerraFluent.Docx.Reporting;

/// <summary>
/// Represents a TerraFluent.Docx.Reporting document that can be published as a DOCX package.
/// </summary>
public sealed class Document
{
    private readonly DocumentContainerDescriptor _container;

    private Document(DocumentContainerDescriptor container)
    {
        _container = container;
    }

    /// <summary>
    /// Creates a new document using the fluent TerraFluent.Docx.Reporting document builder.
    /// </summary>
    /// <param name="configure">Configures document metadata, styles, and pages.</param>
    /// <returns>A document instance that can be published to a file, stream, or byte array.</returns>
    public static Document Create(Action<IDocumentContainer> configure)
    {
        var container = new DocumentContainerDescriptor();
        configure(container);
        return new Document(container);
    }

    /// <summary>
    /// Writes the document to a DOCX file.
    /// </summary>
    /// <param name="filePath">Destination `.docx` file path.</param>
    public void PublishDocx(string filePath)
    {
        using var stream = File.Create(filePath);
        PublishDocx(stream);
    }

    /// <summary>
    /// Writes the document to a DOCX byte array.
    /// </summary>
    /// <returns>The generated DOCX package bytes.</returns>
    public byte[] PublishDocx()
    {
        using var ms = new MemoryStream();
        PublishDocx(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// Writes the document to a stream.
    /// </summary>
    /// <param name="stream">Destination stream. The stream remains open after writing.</param>
    public void PublishDocx(Stream stream)
    {
        OoxmlWriter.Write(_container, stream);
    }
}
