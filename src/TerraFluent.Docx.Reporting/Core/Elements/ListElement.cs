namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class ListElement : IElement
{
    public bool Ordered { get; set; }
    public List<ListItemElement> Items { get; } = [];
    public Dictionary<int, ListLevelMarker> Markers { get; } = [];
}

internal sealed class ListItemElement
{
    public TextElement Text { get; set; } = new();
    public int Level { get; set; }
}

internal sealed class ListLevelMarker
{
    public string Marker { get; set; } = string.Empty;
    public string? FontFamily { get; set; }
}
