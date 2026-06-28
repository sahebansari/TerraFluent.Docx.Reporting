using TerraFluent.Docx.Reporting.Core.Elements;
using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class ColumnDescriptor : IColumnDescriptor
{
    private readonly DocumentTheme _theme;
    private readonly DocumentStyleCatalog _styles;
    internal ColumnElement Element { get; } = new();

    public ColumnDescriptor(DocumentTheme? theme = null, DocumentStyleCatalog? styles = null)
    {
        _theme = theme ?? DocumentTheme.Default;
        _styles = styles ?? new DocumentStyleCatalog();
    }

    public IColumnDescriptor Spacing(float points)
    {
        Element.Spacing = points;
        return this;
    }

    public IContainer Item()
    {
        var container = new ContainerDescriptor(_theme, _styles);
        Element.Items.Add(container.Element);
        return container;
    }
}
