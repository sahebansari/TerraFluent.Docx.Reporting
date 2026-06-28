namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class TableOfContentsElement : IElement
{
    public string Title { get; set; } = "Contents";
    public int MinLevel { get; set; } = 1;
    public int MaxLevel { get; set; } = 3;
}
