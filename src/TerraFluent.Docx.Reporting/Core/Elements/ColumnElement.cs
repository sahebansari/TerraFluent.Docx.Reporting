namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class ColumnElement : IElement
{
    public float Spacing { get; set; }
    public List<ContainerElement> Items { get; } = [];
}
