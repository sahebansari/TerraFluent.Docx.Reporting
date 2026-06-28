namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class ImageElement : IElement
{
    public string? FilePath { get; set; }
    public byte[]? Bytes { get; set; }
    public string FileName { get; set; } = "image.png";
    public float? Width { get; set; }
    public float? Height { get; set; }
    public float? MaxWidth { get; set; }
    public string? AltText { get; set; }
    public string? Alignment { get; set; }
    public TextElement? Caption { get; set; }
    public string? CaptionLabel { get; set; }  // "Figure" for auto-numbered, null for static
    public string? CaptionDescription { get; set; }  // Description text after number
    public int? CaptionFigureNumber { get; set; }  // Auto-assigned figure number
    public string WrapMode { get; set; } = "inline";
    public string HorizontalRelativeFrom { get; set; } = "margin";
    public string VerticalRelativeFrom { get; set; } = "paragraph";
    public string? HorizontalAlignment { get; set; }
    public string? VerticalAlignment { get; set; }
    public float? HorizontalOffset { get; set; }
    public float? VerticalOffset { get; set; }
    public float MarginTop { get; set; }
    public float MarginRight { get; set; }
    public float MarginBottom { get; set; }
    public float MarginLeft { get; set; }
    public float? BorderWidth { get; set; }
    public string BorderColor { get; set; } = "000000";
    public bool Rounded { get; set; }
    public float CropLeftPercent { get; set; }
    public float CropTopPercent { get; set; }
    public float CropRightPercent { get; set; }
    public float CropBottomPercent { get; set; }
}
