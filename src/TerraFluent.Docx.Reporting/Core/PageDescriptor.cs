using TerraFluent.Docx.Reporting.Core.Elements;
using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class PageDescriptor : IPageDescriptor
{
    private readonly DocumentTheme _theme;
    private readonly DocumentStyleCatalog _styles;
    public PageSize Size { get; private set; } = PageSize.A4;
    public float MarginTop    { get; private set; } = Unit.Centimetre(2.54f);
    public float MarginRight  { get; private set; } = Unit.Centimetre(2.54f);
    public float MarginBottom { get; private set; } = Unit.Centimetre(2.54f);
    public float MarginLeft   { get; private set; } = Unit.Centimetre(2.54f);

    internal ContainerDescriptor HeaderContainer { get; private set; }
    internal ContainerDescriptor FirstPageHeaderContainer { get; private set; }
    internal ContainerDescriptor EvenPageHeaderContainer { get; private set; }
    internal ContainerDescriptor ContentContainer { get; private set; }
    internal ContainerDescriptor FooterContainer { get; private set; }
    internal ContainerDescriptor FirstPageFooterContainer { get; private set; }
    internal ContainerDescriptor EvenPageFooterContainer { get; private set; }
    internal TextElement DefaultTextStyle { get; private set; } = new();
    internal int? PageNumberStartValue { get; private set; }
    internal string? PageNumberFormat { get; private set; }
    internal string? BackgroundColor { get; private set; }
    internal string? WatermarkText { get; private set; }
    internal string WatermarkColor { get; private set; } = "D9D9D9";
    internal float WatermarkFontSize { get; private set; } = 54;
    internal int ColumnCount { get; private set; } = 1;
    internal float ColumnSpacing { get; private set; } = 36;
    internal bool ColumnSeparatorLine { get; private set; }

    public PageDescriptor(DocumentTheme theme, DocumentStyleCatalog styles)
    {
        _theme = theme;
        _styles = styles;
        HeaderContainer = new ContainerDescriptor(theme, styles);
        FirstPageHeaderContainer = new ContainerDescriptor(theme, styles);
        EvenPageHeaderContainer = new ContainerDescriptor(theme, styles);
        ContentContainer = new ContainerDescriptor(theme, styles);
        FooterContainer = new ContainerDescriptor(theme, styles);
        FirstPageFooterContainer = new ContainerDescriptor(theme, styles);
        EvenPageFooterContainer = new ContainerDescriptor(theme, styles);
        ApplyTheme(theme);
    }

    internal void ApplyTheme(DocumentTheme theme)
    {
        var descriptor = new TextDescriptor(string.Empty, theme);
        descriptor.FontFamily(theme.DefaultFontFamily)
            .FontSize(theme.DefaultFontSize)
            .FontColor(theme.DefaultTextColor);
        DefaultTextStyle = descriptor.Element;
    }

    IPageDescriptor IPageDescriptor.Size(PageSize size)
    {
        Size = new PageSize(
            Guard.PositiveFinite(size.Width, nameof(size.Width)),
            Guard.PositiveFinite(size.Height, nameof(size.Height)));
        return this;
    }

    IPageDescriptor IPageDescriptor.Size(float w, float h)
    {
        Size = new PageSize(
            Guard.PositiveFinite(w, nameof(w)),
            Guard.PositiveFinite(h, nameof(h)));
        return this;
    }
    IPageDescriptor IPageDescriptor.Landscape()                       { Size = Size.Landscape(); return this; }
    IPageDescriptor IPageDescriptor.Portrait()                        { Size = new PageSize(Math.Min(Size.Width, Size.Height), Math.Max(Size.Width, Size.Height)); return this; }

    IPageDescriptor IPageDescriptor.Margin(float all)
    {
        Guard.NonNegativeFinite(all, nameof(all));
        MarginTop = MarginRight = MarginBottom = MarginLeft = all;
        return this;
    }

    IPageDescriptor IPageDescriptor.Margin(float v, float h)
    {
        Guard.NonNegativeFinite(v, nameof(v));
        Guard.NonNegativeFinite(h, nameof(h));
        MarginTop = MarginBottom = v;
        MarginRight = MarginLeft = h;
        return this;
    }

    IPageDescriptor IPageDescriptor.Margin(float top, float right, float bottom, float left)
    {
        Guard.NonNegativeFinite(top, nameof(top));
        Guard.NonNegativeFinite(right, nameof(right));
        Guard.NonNegativeFinite(bottom, nameof(bottom));
        Guard.NonNegativeFinite(left, nameof(left));
        MarginTop = top; MarginRight = right; MarginBottom = bottom; MarginLeft = left;
        return this;
    }

    IPageDescriptor IPageDescriptor.MarginTop(float p)    { MarginTop = Guard.NonNegativeFinite(p, nameof(p)); return this; }
    IPageDescriptor IPageDescriptor.MarginRight(float p)  { MarginRight = Guard.NonNegativeFinite(p, nameof(p)); return this; }
    IPageDescriptor IPageDescriptor.MarginBottom(float p) { MarginBottom = Guard.NonNegativeFinite(p, nameof(p)); return this; }
    IPageDescriptor IPageDescriptor.MarginLeft(float p)   { MarginLeft = Guard.NonNegativeFinite(p, nameof(p)); return this; }

    IPageDescriptor IPageDescriptor.DefaultTextStyle(Action<ITextDescriptor> configure)
    {
        var descriptor = new TextDescriptor(string.Empty, _theme);
        Guard.NotNull(configure, nameof(configure));
        configure(descriptor);
        DefaultTextStyle = descriptor.Element;
        return this;
    }

    IPageDescriptor IPageDescriptor.PageNumberStart(int value)
    {
        PageNumberStartValue = Guard.InRange(value, 1, int.MaxValue, nameof(value));
        return this;
    }

    IPageDescriptor IPageDescriptor.PageNumberFormat(string format)
    {
        PageNumberFormat = Guard.NotWhiteSpace(format, nameof(format));
        return this;
    }

    IPageDescriptor IPageDescriptor.Background(string hexColor)
    {
        BackgroundColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    IPageDescriptor IPageDescriptor.Watermark(string text, string hexColor, float fontSize)
    {
        WatermarkText = Guard.NotWhiteSpace(text, nameof(text));
        WatermarkColor = HexColor.Validate(hexColor, nameof(hexColor));
        WatermarkFontSize = Guard.PositiveFinite(fontSize, nameof(fontSize));
        return this;
    }

    IPageDescriptor IPageDescriptor.Columns(int count, float spacingPoints, bool separatorLine)
    {
        ColumnCount = Guard.InRange(count, 1, 45, nameof(count));
        ColumnSpacing = Guard.NonNegativeFinite(spacingPoints, nameof(spacingPoints));
        ColumnSeparatorLine = separatorLine;
        return this;
    }

    IPageDescriptor IPageDescriptor.SingleColumn()
    {
        ColumnCount = 1;
        ColumnSpacing = 36;
        ColumnSeparatorLine = false;
        return this;
    }

    IContainer IPageDescriptor.Header()          => HeaderContainer;
    IContainer IPageDescriptor.OddPageHeader()   => HeaderContainer;
    IContainer IPageDescriptor.FirstPageHeader() => FirstPageHeaderContainer;
    IContainer IPageDescriptor.EvenPageHeader()  => EvenPageHeaderContainer;
    IContainer IPageDescriptor.Content()         => ContentContainer;
    IContainer IPageDescriptor.Footer()          => FooterContainer;
    IContainer IPageDescriptor.OddPageFooter()   => FooterContainer;
    IContainer IPageDescriptor.FirstPageFooter() => FirstPageFooterContainer;
    IContainer IPageDescriptor.EvenPageFooter()  => EvenPageFooterContainer;
}
