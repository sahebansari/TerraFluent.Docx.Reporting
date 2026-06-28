using TerraFluent.Docx.Reporting.Core.Elements;
using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class RowDescriptor : IRowDescriptor
{
    private readonly DocumentTheme _theme;
    private readonly DocumentStyleCatalog _styles;
    internal RowElement Element { get; } = new();

    public RowDescriptor(DocumentTheme? theme = null, DocumentStyleCatalog? styles = null)
    {
        _theme = theme ?? DocumentTheme.Default;
        _styles = styles ?? new DocumentStyleCatalog();
    }

    public IRowDescriptor Spacing(float points)
    {
        Element.Spacing = points;
        return this;
    }

    public IContainer RelativeItem(float size = 1)
    {
        var container = new ContainerDescriptor(_theme, _styles);
        Element.Cells.Add(new RowCell { Mode = RowCell.SizingMode.Relative, Size = size, Container = container.Element });
        return container;
    }

    public IContainer AutoItem()
    {
        var container = new ContainerDescriptor(_theme, _styles);
        Element.Cells.Add(new RowCell { Mode = RowCell.SizingMode.Auto, Container = container.Element });
        return container;
    }

    public IContainer ConstantItem(float widthPoints)
    {
        var container = new ContainerDescriptor(_theme, _styles);
        Element.Cells.Add(new RowCell { Mode = RowCell.SizingMode.Constant, Size = widthPoints, Container = container.Element });
        return container;
    }
}
