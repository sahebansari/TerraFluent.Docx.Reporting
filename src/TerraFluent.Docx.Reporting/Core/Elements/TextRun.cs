namespace TerraFluent.Docx.Reporting.Core.Elements;

internal enum TextRunKind { Text, Tab, Hyperlink, Field, Footnote, Endnote, PageNumber, PageCount }

internal sealed class TextRun
{
    public TextRunKind Kind { get; set; } = TextRunKind.Text;
    public string Text { get; set; } = string.Empty;
    public string? Url { get; set; }
    public string? FieldInstruction { get; set; }
    public string? FieldCachedText { get; set; }
    public TextElement? NoteText { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public bool? Underline { get; set; }
    public bool? Strikethrough { get; set; }
    public float? FontSize { get; set; }
    public string? FontColor { get; set; }
    public string? FontFamily { get; set; }
    public string? VerticalAlignment { get; set; }
    public string? HighlightColor { get; set; }
}
