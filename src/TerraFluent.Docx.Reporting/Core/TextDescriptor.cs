using TerraFluent.Docx.Reporting.Core.Elements;
using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class TextDescriptor : ITextDescriptor
{
    private readonly DocumentTheme _theme;
    private readonly bool _isNoteBody;
    internal TextElement Element { get; } = new();

    private TextRun _currentRun;

    public TextDescriptor(string text, DocumentTheme? theme = null, bool isNoteBody = false)
    {
        _theme = theme ?? DocumentTheme.Default;
        _isNoteBody = isNoteBody;
        _currentRun = new TextRun { Text = text };
        Element.Runs.Add(_currentRun);
    }

    public ITextDescriptor Bold(bool value = true)          { _currentRun.Bold = value; return this; }
    public ITextDescriptor Italic(bool value = true)        { _currentRun.Italic = value; return this; }
    public ITextDescriptor Underline(bool value = true)     { _currentRun.Underline = value; return this; }
    public ITextDescriptor Strikethrough(bool value = true) { _currentRun.Strikethrough = value; return this; }
    public ITextDescriptor Superscript(bool value = true)   { SetVerticalAlignment(_currentRun, "superscript", value); return this; }
    public ITextDescriptor Subscript(bool value = true)     { SetVerticalAlignment(_currentRun, "subscript", value); return this; }
    public ITextDescriptor FontSize(float size)             { _currentRun.FontSize = size; return this; }
    public ITextDescriptor FontColor(string hexColor)       { _currentRun.FontColor = HexColor.Validate(hexColor, nameof(hexColor)); return this; }
    public ITextDescriptor FontFamily(string name)          { _currentRun.FontFamily = name; return this; }
    public ITextDescriptor Highlight(HighlightColor color = HighlightColor.Yellow) { _currentRun.HighlightColor = color.ToOoxmlValue(); return this; }
    public ITextDescriptor Style(string name)               { Element.StyleId = DocumentStyleCatalog.StyleId(name); return this; }

    public ITextDescriptor AlignLeft()   { Element.Alignment = "left";   return this; }
    public ITextDescriptor AlignCenter() { Element.Alignment = "center"; return this; }
    public ITextDescriptor AlignRight()  { Element.Alignment = "right";  return this; }
    public ITextDescriptor Justify()     { Element.Alignment = "both";   return this; }

    public ITextDescriptor LineHeight(float multiplier) { Element.LineHeight = multiplier; return this; }
    public ITextDescriptor SpacingBefore(float points)  { Element.SpacingBefore = points;  return this; }
    public ITextDescriptor SpacingAfter(float points)   { Element.SpacingAfter = points;   return this; }
    public ITextDescriptor LeftIndent(float points)     { Element.LeftIndent = points;     return this; }
    public ITextDescriptor RightIndent(float points)    { Element.RightIndent = points;    return this; }
    public ITextDescriptor FirstLineIndent(float points) { Element.FirstLineIndent = points; return this; }
    public ITextDescriptor HangingIndent(float points)  { Element.HangingIndent = points;  return this; }
    public ITextDescriptor KeepWithNext(bool value = true)      { Element.KeepWithNext = value;      return this; }
    public ITextDescriptor KeepLinesTogether(bool value = true) { Element.KeepLinesTogether = value; return this; }
    public ITextDescriptor PageBreakBefore(bool value = true)   { Element.PageBreakBefore = value;   return this; }
    public ITextDescriptor Shading(string hexColor)             { Element.ShadingColor = HexColor.Validate(hexColor, nameof(hexColor)); return this; }
    public ITextDescriptor Border(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f)
    {
        BorderTop(widthPoints, hexColor, spacePoints);
        BorderRight(widthPoints, hexColor, spacePoints);
        BorderBottom(widthPoints, hexColor, spacePoints);
        BorderLeft(widthPoints, hexColor, spacePoints);
        return this;
    }

    public ITextDescriptor BorderTop(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f)
    {
        Element.TopBorder = CreateBorder(widthPoints, hexColor, spacePoints);
        return this;
    }

    public ITextDescriptor BorderRight(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f)
    {
        Element.RightBorder = CreateBorder(widthPoints, hexColor, spacePoints);
        return this;
    }

    public ITextDescriptor BorderBottom(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f)
    {
        Element.BottomBorder = CreateBorder(widthPoints, hexColor, spacePoints);
        return this;
    }

    public ITextDescriptor BorderLeft(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f)
    {
        Element.LeftBorder = CreateBorder(widthPoints, hexColor, spacePoints);
        return this;
    }

    public ITextDescriptor TabStop(float positionPoints, TabStopAlignment alignment = TabStopAlignment.Left)
    {
        Element.TabStops.Add(new TabStopElement
        {
            Position = positionPoints,
            Alignment = alignment
        });
        return this;
    }

    public ITextDescriptor Span(string text, Action<ITextDescriptor>? configure = null)
    {
        var run = CreateRun(text);
        if (configure != null) new SpanDescriptor(run).Apply(configure);
        Element.Runs.Add(run);
        return new SpanDescriptor(run, this);
    }

    public ITextDescriptor Tab()
    {
        var run = CreateRun(string.Empty);
        run.Kind = TextRunKind.Tab;
        Element.Runs.Add(run);
        return new SpanDescriptor(run, this);
    }

    public ITextDescriptor Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null)
    {
        var run = CreateRun(text);
        run.Kind = TextRunKind.Hyperlink;
        run.Url = url;
        run.Underline = true;
        run.FontColor ??= _theme.HyperlinkColor;
        if (configure != null) new SpanDescriptor(run).Apply(configure);
        Element.Runs.Add(run);
        return new SpanDescriptor(run, this);
    }

    public ITextDescriptor CrossReference(string bookmarkName, string? fallbackText = null, Action<ITextDescriptor>? configure = null)
    {
        var run = CreateRun(string.Empty);
        run.Kind = TextRunKind.Field;
        run.FieldInstruction = $"REF {BookmarkName(bookmarkName)} \\h";
        run.FieldCachedText = fallbackText ?? bookmarkName;
        if (configure != null) new SpanDescriptor(run).Apply(configure);
        Element.Runs.Add(run);
        return new SpanDescriptor(run, this);
    }

    public ITextDescriptor Footnote(string text, Action<ITextDescriptor>? configure = null)
        => AddNote(TextRunKind.Footnote, text, configure);

    public ITextDescriptor Endnote(string text, Action<ITextDescriptor>? configure = null)
        => AddNote(TextRunKind.Endnote, text, configure);

    private ITextDescriptor AddNote(TextRunKind kind, string text, Action<ITextDescriptor>? configure)
    {
        if (_isNoteBody)
            throw new InvalidOperationException("Footnotes and endnotes cannot contain nested footnote or endnote references.");

        var noteDescriptor = new TextDescriptor(text, _theme, isNoteBody: true);
        configure?.Invoke(noteDescriptor);
        var run = CreateRun(string.Empty);
        run.Kind = kind;
        run.NoteText = noteDescriptor.Element;
        Element.Runs.Add(run);
        return new SpanDescriptor(run, this);
    }

    public ITextDescriptor CurrentPageNumber(Action<ITextDescriptor>? configure = null)
        => AddFieldRun(TextRunKind.PageNumber, configure);

    public ITextDescriptor TotalPages(Action<ITextDescriptor>? configure = null)
        => AddFieldRun(TextRunKind.PageCount, configure);

    private ITextDescriptor AddFieldRun(TextRunKind kind, Action<ITextDescriptor>? configure)
    {
        var run = CreateRun(string.Empty);
        run.Kind = kind;
        if (configure != null) new SpanDescriptor(run).Apply(configure);
        Element.Runs.Add(run);
        return new SpanDescriptor(run, this);
    }

    // New runs start unset; BuildRPr falls back to the paragraph's first run (_currentRun)
    // for any property not explicitly overridden on the span itself.
    private static TextRun CreateRun(string text) => new() { Text = text };

    internal static void SetVerticalAlignment(TextRun run, string alignment, bool value)
    {
        if (value)
            run.VerticalAlignment = alignment;
        else if (run.VerticalAlignment == alignment)
            run.VerticalAlignment = null;
    }

    internal static string BookmarkName(string name)
    {
        var normalized = new string(name.Select(ch => char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_').ToArray());
        if (string.IsNullOrWhiteSpace(normalized))
            normalized = "Bookmark";
        if (char.IsDigit(normalized[0]))
            normalized = "_" + normalized;
        return normalized;
    }

    private static ParagraphBorder CreateBorder(float widthPoints, string hexColor, float spacePoints) => new()
    {
        Width = Math.Max(0, widthPoints),
        Color = HexColor.Validate(hexColor, nameof(hexColor)),
        Space = Math.Max(0, spacePoints)
    };
}

internal sealed class SpanDescriptor : ITextDescriptor
{
    private readonly TextRun _run;
    private readonly TextDescriptor? _parent;

    public SpanDescriptor(TextRun run, TextDescriptor? parent = null)
    {
        _run = run;
        _parent = parent;
    }

    internal void Apply(Action<ITextDescriptor> configure) => configure(this);

    public ITextDescriptor Bold(bool value = true)          { _run.Bold = value;         return this; }
    public ITextDescriptor Italic(bool value = true)        { _run.Italic = value;       return this; }
    public ITextDescriptor Underline(bool value = true)     { _run.Underline = value;    return this; }
    public ITextDescriptor Strikethrough(bool value = true) { _run.Strikethrough = value; return this; }
    public ITextDescriptor Superscript(bool value = true)   { TextDescriptor.SetVerticalAlignment(_run, "superscript", value); return this; }
    public ITextDescriptor Subscript(bool value = true)     { TextDescriptor.SetVerticalAlignment(_run, "subscript", value); return this; }
    public ITextDescriptor FontSize(float size)             { _run.FontSize = size;       return this; }
    public ITextDescriptor FontColor(string hexColor)       { _run.FontColor = HexColor.Validate(hexColor, nameof(hexColor)); return this; }
    public ITextDescriptor FontFamily(string name)          { _run.FontFamily = name;     return this; }
    public ITextDescriptor Highlight(HighlightColor color = HighlightColor.Yellow) { _run.HighlightColor = color.ToOoxmlValue(); return this; }
    public ITextDescriptor Style(string name)                              => _parent?.Style(name) ?? this;

    public ITextDescriptor AlignLeft()                                     => _parent?.AlignLeft() ?? this;
    public ITextDescriptor AlignCenter()                                   => _parent?.AlignCenter() ?? this;
    public ITextDescriptor AlignRight()                                    => _parent?.AlignRight() ?? this;
    public ITextDescriptor Justify()                                       => _parent?.Justify() ?? this;
    public ITextDescriptor LineHeight(float multiplier)                    => _parent?.LineHeight(multiplier) ?? this;
    public ITextDescriptor SpacingBefore(float points)                     => _parent?.SpacingBefore(points) ?? this;
    public ITextDescriptor SpacingAfter(float points)                      => _parent?.SpacingAfter(points) ?? this;
    public ITextDescriptor LeftIndent(float points)                        => _parent?.LeftIndent(points) ?? this;
    public ITextDescriptor RightIndent(float points)                       => _parent?.RightIndent(points) ?? this;
    public ITextDescriptor FirstLineIndent(float points)                   => _parent?.FirstLineIndent(points) ?? this;
    public ITextDescriptor HangingIndent(float points)                     => _parent?.HangingIndent(points) ?? this;
    public ITextDescriptor KeepWithNext(bool value = true)                 => _parent?.KeepWithNext(value) ?? this;
    public ITextDescriptor KeepLinesTogether(bool value = true)            => _parent?.KeepLinesTogether(value) ?? this;
    public ITextDescriptor PageBreakBefore(bool value = true)              => _parent?.PageBreakBefore(value) ?? this;
    public ITextDescriptor Shading(string hexColor)                        => _parent?.Shading(hexColor) ?? this;
    public ITextDescriptor Border(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f) => _parent?.Border(widthPoints, hexColor, spacePoints) ?? this;
    public ITextDescriptor BorderTop(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f) => _parent?.BorderTop(widthPoints, hexColor, spacePoints) ?? this;
    public ITextDescriptor BorderRight(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f) => _parent?.BorderRight(widthPoints, hexColor, spacePoints) ?? this;
    public ITextDescriptor BorderBottom(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f) => _parent?.BorderBottom(widthPoints, hexColor, spacePoints) ?? this;
    public ITextDescriptor BorderLeft(float widthPoints = 1f, string hexColor = "000000", float spacePoints = 4f) => _parent?.BorderLeft(widthPoints, hexColor, spacePoints) ?? this;
    public ITextDescriptor TabStop(float positionPoints, TabStopAlignment alignment = TabStopAlignment.Left) => _parent?.TabStop(positionPoints, alignment) ?? this;
    public ITextDescriptor Span(string text, Action<ITextDescriptor>? c)   => _parent?.Span(text, c) ?? this;
    public ITextDescriptor Tab()                                           => _parent?.Tab() ?? this;
    public ITextDescriptor Hyperlink(string text, string url, Action<ITextDescriptor>? c) => _parent?.Hyperlink(text, url, c) ?? this;
    public ITextDescriptor CrossReference(string bookmarkName, string? fallbackText = null, Action<ITextDescriptor>? c = null) => _parent?.CrossReference(bookmarkName, fallbackText, c) ?? this;
    public ITextDescriptor Footnote(string text, Action<ITextDescriptor>? c = null) => _parent?.Footnote(text, c) ?? this;
    public ITextDescriptor Endnote(string text, Action<ITextDescriptor>? c = null) => _parent?.Endnote(text, c) ?? this;
    public ITextDescriptor CurrentPageNumber(Action<ITextDescriptor>? c)   => _parent?.CurrentPageNumber(c) ?? this;
    public ITextDescriptor TotalPages(Action<ITextDescriptor>? c)          => _parent?.TotalPages(c) ?? this;
}
