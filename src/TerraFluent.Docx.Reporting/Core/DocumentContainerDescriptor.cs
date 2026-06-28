using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class DocumentContainerDescriptor : IDocumentContainer
{
    internal string? Title    { get; private set; }
    internal string? Author   { get; private set; }
    internal string? Subject  { get; private set; }
    internal string? Keywords { get; private set; }
    internal string? Creator  { get; private set; }
    internal List<PageDescriptor> Pages { get; } = [];
    internal DocumentTheme ThemeSettings { get; private set; } = DocumentTheme.Default;
    internal DocumentStyleCatalog Styles { get; } = new();

    public IDocumentContainer Theme(DocumentTheme theme)
    {
        // Mutate the existing instance in place (rather than replacing it) so that
        // descriptors created before this call - which hold a reference to this same
        // instance - pick up the new values for any elements added afterward.
        ThemeSettings.CopyFrom(theme);
        foreach (var page in Pages)
            page.ApplyTheme(ThemeSettings);
        return this;
    }

    public IDocumentContainer Theme(Action<IDocumentThemeDescriptor> configure)
    {
        var theme = ThemeSettings.Clone();
        configure(new DocumentThemeDescriptor(theme));
        return Theme(theme);
    }

    public IDocumentContainer ParagraphStyle(string name, Action<ITextDescriptor> configure)
    {
        var descriptor = new TextDescriptor(string.Empty, ThemeSettings);
        configure(descriptor);
        Styles.AddParagraphStyle(name, descriptor.Element);
        return this;
    }

    public IDocumentContainer TableStyle(string name, Action<ITableDescriptor> configure)
    {
        var descriptor = new TableDescriptor(ThemeSettings, Styles);
        configure(descriptor);
        Styles.AddTableStyle(name, descriptor.Element);
        return this;
    }

    public IDocumentContainer Page(Action<IPageDescriptor> configure)
    {
        var page = new PageDescriptor(ThemeSettings, Styles);
        configure(page);
        Pages.Add(page);
        return this;
    }

    public IDocumentContainer MetadataTitle(string title)       { Title = title; return this; }
    public IDocumentContainer MetadataAuthor(string author)     { Author = author; return this; }
    public IDocumentContainer MetadataSubject(string subject)   { Subject = subject; return this; }
    public IDocumentContainer MetadataKeywords(string keywords) { Keywords = keywords; return this; }
    public IDocumentContainer MetadataCreator(string creator)   { Creator = creator; return this; }

    public IDocumentContainer Compose(IDocument document)
    {
        document.Compose(this);
        return this;
    }
}

internal sealed class DocumentThemeDescriptor : IDocumentThemeDescriptor
{
    private readonly DocumentTheme _theme;

    public DocumentThemeDescriptor(DocumentTheme theme)
    {
        _theme = theme;
    }

    public IDocumentThemeDescriptor DefaultFont(string family, float size = 11)
    {
        _theme.DefaultFontFamily = family;
        _theme.DefaultFontSize = size;
        return this;
    }

    public IDocumentThemeDescriptor DefaultFontFamily(string family)
    {
        _theme.DefaultFontFamily = family;
        return this;
    }

    public IDocumentThemeDescriptor DefaultFontSize(float size)
    {
        _theme.DefaultFontSize = size;
        return this;
    }

    public IDocumentThemeDescriptor DefaultTextColor(string hexColor)
    {
        _theme.DefaultTextColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public IDocumentThemeDescriptor HeadingColor(string hexColor)
    {
        _theme.HeadingColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public IDocumentThemeDescriptor AccentColor(string hexColor)
    {
        var color = HexColor.Validate(hexColor, nameof(hexColor));
        _theme.AccentColor = color;
        _theme.HeadingColor = color;
        _theme.TableHeaderBackgroundColor ??= color;
        return this;
    }

    public IDocumentThemeDescriptor HyperlinkColor(string hexColor)
    {
        _theme.HyperlinkColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public IDocumentThemeDescriptor TableHeaderBackground(string hexColor)
    {
        _theme.TableHeaderBackgroundColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public IDocumentThemeDescriptor TableAlternateRowBackground(string hexColor)
    {
        _theme.TableAlternateRowBackgroundColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public IDocumentThemeDescriptor TableBorder(float width, string hexColor)
    {
        _theme.TableBorderWidth = width;
        _theme.TableBorderColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public IDocumentThemeDescriptor TableCellPadding(float points)
    {
        _theme.TableCellPaddingVertical = points;
        _theme.TableCellPaddingHorizontal = points;
        return this;
    }

    public IDocumentThemeDescriptor TableCellPadding(float verticalPoints, float horizontalPoints)
    {
        _theme.TableCellPaddingVertical = verticalPoints;
        _theme.TableCellPaddingHorizontal = horizontalPoints;
        return this;
    }

    public IDocumentThemeDescriptor TableRowMinHeight(float points)
    {
        _theme.TableRowMinHeight = points;
        return this;
    }

    public IDocumentThemeDescriptor TableHeaderRowMinHeight(float points)
    {
        _theme.TableHeaderRowMinHeight = points;
        return this;
    }
}
