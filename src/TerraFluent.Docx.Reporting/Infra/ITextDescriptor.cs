namespace TerraFluent.Docx.Reporting.Infra;

public interface ITextDescriptor
{
    ITextDescriptor Bold(bool value = true);
    ITextDescriptor Italic(bool value = true);
    ITextDescriptor Underline(bool value = true);
    ITextDescriptor Strikethrough(bool value = true);
    ITextDescriptor Superscript(bool value = true);
    ITextDescriptor Subscript(bool value = true);
    ITextDescriptor FontSize(float size);
    ITextDescriptor FontColor(string hexColor);
    ITextDescriptor FontFamily(string name);
    ITextDescriptor Highlight(HighlightColor color = HighlightColor.Yellow);
    ITextDescriptor Style(string name);
    ITextDescriptor AlignLeft();
    ITextDescriptor AlignCenter();
    ITextDescriptor AlignRight();
    ITextDescriptor Justify();
    ITextDescriptor LineHeight(float multiplier);
    ITextDescriptor SpacingBefore(float points);
    ITextDescriptor SpacingAfter(float points);
    ITextDescriptor LeftIndent(float points);
    ITextDescriptor RightIndent(float points);
    ITextDescriptor FirstLineIndent(float points);
    ITextDescriptor HangingIndent(float points);
    ITextDescriptor KeepWithNext(bool value = true);
    ITextDescriptor KeepLinesTogether(bool value = true);
    ITextDescriptor PageBreakBefore(bool value = true);
    ITextDescriptor Shading(string hexColor);
    ITextDescriptor Border(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);
    ITextDescriptor BorderTop(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);
    ITextDescriptor BorderRight(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);
    ITextDescriptor BorderBottom(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);
    ITextDescriptor BorderLeft(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f);
    ITextDescriptor TabStop(float positionPoints, TabStopAlignment alignment = TabStopAlignment.Left);
    ITextDescriptor Span(string text, Action<ITextDescriptor>? configure = null);
    ITextDescriptor Tab();
    ITextDescriptor Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null);
    ITextDescriptor CrossReference(string bookmarkName, string? fallbackText = null, Action<ITextDescriptor>? configure = null);
    ITextDescriptor Footnote(string text, Action<ITextDescriptor>? configure = null);
    ITextDescriptor Endnote(string text, Action<ITextDescriptor>? configure = null);
    ITextDescriptor CurrentPageNumber(Action<ITextDescriptor>? configure = null);
    ITextDescriptor TotalPages(Action<ITextDescriptor>? configure = null);
}
