namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class BarcodeElement : IElement
{
    public string Value { get; set; } = "";
    public float? Width { get; set; }
    public float? Height { get; set; }
    public float? MaxWidth { get; set; }
    public string? AltText { get; set; }
    public string? Alignment { get; set; }
    public bool ShowText { get; set; } = true;
    public string BarColor { get; set; } = "000000";
    public TextElement? Caption { get; set; }
}
