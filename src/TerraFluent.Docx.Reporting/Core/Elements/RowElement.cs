namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class RowElement : IElement
{
    public float Spacing { get; set; }
    public List<RowCell> Cells { get; } = [];
}

internal sealed class RowCell
{
    public enum SizingMode { Relative, Auto, Constant }
    public SizingMode Mode { get; set; } = SizingMode.Relative;
    public float Size { get; set; } = 1;
    public ContainerElement Container { get; set; } = new();
}
