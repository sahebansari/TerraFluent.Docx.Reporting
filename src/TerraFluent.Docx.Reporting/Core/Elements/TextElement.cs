namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class TextElement : IElement
{
    public List<TextRun> Runs { get; } = [];
    public string Alignment { get; set; } = "left";
    public float? LineHeight { get; set; }
    public float SpacingBefore { get; set; }
    public float SpacingAfter { get; set; }
    public float? LeftIndent { get; set; }
    public float? RightIndent { get; set; }
    public float? FirstLineIndent { get; set; }
    public float? HangingIndent { get; set; }
    public bool KeepWithNext { get; set; }
    public bool KeepLinesTogether { get; set; }
    public bool PageBreakBefore { get; set; }
    public string? StyleId { get; set; }
    public string? BookmarkName { get; set; }
    public string? ShadingColor { get; set; }
    public ParagraphBorder? TopBorder { get; set; }
    public ParagraphBorder? RightBorder { get; set; }
    public ParagraphBorder? BottomBorder { get; set; }
    public ParagraphBorder? LeftBorder { get; set; }
    public List<TabStopElement> TabStops { get; } = [];
}

internal sealed class ParagraphBorder
{
    public float Width { get; set; }
    public string Color { get; set; } = "000000";
    public float Space { get; set; }
}

internal sealed class TabStopElement
{
    public float Position { get; set; }
    public TabStopAlignment Alignment { get; set; }
}
