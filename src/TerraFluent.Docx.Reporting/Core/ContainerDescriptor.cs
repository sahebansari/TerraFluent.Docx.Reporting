using TerraFluent.Docx.Reporting.Core.Elements;
using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class ContainerDescriptor : IContainer
{
    private readonly DocumentTheme _theme;
    private readonly DocumentStyleCatalog _styles;
    internal ContainerElement Element { get; } = new();

    public ContainerDescriptor(DocumentTheme? theme = null, DocumentStyleCatalog? styles = null)
    {
        _theme = theme ?? DocumentTheme.Default;
        _styles = styles ?? new DocumentStyleCatalog();
    }

    public IContainer H1(string text) => Heading(text, "Heading1");
    public IContainer H2(string text) => Heading(text, "Heading2");
    public IContainer H3(string text) => Heading(text, "Heading3");
    public IContainer H4(string text) => Heading(text, "Heading4");
    public IContainer H5(string text) => Heading(text, "Heading5");
    public IContainer H6(string text) => Heading(text, "Heading6");

    private IContainer Heading(string text, string styleId)
    {
        var td = new TextDescriptor(text, _theme);
        td.Element.StyleId = styleId;
        Element.Elements.Add(td.Element);
        return this;
    }

    public IContainer Column(Action<IColumnDescriptor> configure)
    {
        var col = new ColumnDescriptor(_theme, _styles);
        Guard.NotNull(configure, nameof(configure));
        configure(col);
        Element.Elements.Add(col.Element);
        return this;
    }

    public IContainer Row(Action<IRowDescriptor> configure)
    {
        var row = new RowDescriptor(_theme, _styles);
        Guard.NotNull(configure, nameof(configure));
        configure(row);
        Element.Elements.Add(row.Element);
        return this;
    }

    public IContainer Text(string text, Action<ITextDescriptor>? configure = null)
    {
        var td = new TextDescriptor(text, _theme);
        configure?.Invoke(td);
        Element.Elements.Add(td.Element);
        return this;
    }

    public IContainer Text(Action<ITextDescriptor> configure)
    {
        var td = new TextDescriptor(string.Empty, _theme);
        Guard.NotNull(configure, nameof(configure));
        configure(td);
        Element.Elements.Add(td.Element);
        return this;
    }

    public IContainer Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null)
    {
        var td = new TextDescriptor(string.Empty, _theme);
        td.Hyperlink(text, url, configure);
        Element.Elements.Add(td.Element);
        return this;
    }

    public IContainer Bookmark(string name)
    {
        Element.Elements.Add(new BookmarkElement { Name = name });
        return this;
    }

    public IContainer Bookmark(string name, string text, Action<ITextDescriptor>? configure = null)
    {
        var td = new TextDescriptor(text, _theme);
        td.Element.BookmarkName = name;
        configure?.Invoke(td);
        Element.Elements.Add(td.Element);
        return this;
    }

    public IContainer TableOfContents(string title = "Contents", int minLevel = 1, int maxLevel = 3)
    {
        Element.Elements.Add(new TableOfContentsElement
        {
            Title = title,
            MinLevel = MathCompat.Clamp(minLevel, 1, 9),
            MaxLevel = MathCompat.Clamp(maxLevel, 1, 9)
        });
        return this;
    }

    public IContainer BulletList(Action<IListDescriptor> configure)
        => List(ordered: false, configure);

    public IContainer NumberedList(Action<IListDescriptor> configure)
        => List(ordered: true, configure);

    private IContainer List(bool ordered, Action<IListDescriptor> configure)
    {
        var descriptor = new ListDescriptor(ordered, _theme);
        Guard.NotNull(configure, nameof(configure));
        configure(descriptor);
        Element.Elements.Add(descriptor.Element);
        return this;
    }

    public IContainer Table(Action<ITableDescriptor> configure)
    {
        var td = new TableDescriptor(_theme, _styles);
        Guard.NotNull(configure, nameof(configure));
        configure(td);
        Element.Elements.Add(td.Element);
        return this;
    }

    public IContainer Chart(Action<IChartDescriptor> configure)
    {
        var descriptor = new ChartDescriptor(_theme);
        Guard.NotNull(configure, nameof(configure));
        configure(descriptor);
        Element.Elements.Add(descriptor.Element);
        return this;
    }

    public IContainer Image(string filePath, float? width = null)
    {
        Guard.NotWhiteSpace(filePath, nameof(filePath));
        if (width.HasValue)
            Guard.PositiveFinite(width.Value, nameof(width));

        Element.Elements.Add(new ImageElement
        {
            FilePath = filePath,
            FileName = Path.GetFileName(filePath),
            Width = width
        });
        return this;
    }

    public IContainer Image(string filePath, Action<IImageDescriptor> configure)
    {
        Guard.NotWhiteSpace(filePath, nameof(filePath));
        Guard.NotNull(configure, nameof(configure));

        var image = new ImageElement
        {
            FilePath = filePath,
            FileName = Path.GetFileName(filePath)
        };
        configure(new ImageDescriptor(image, _theme));
        Element.Elements.Add(image);
        return this;
    }

    public IContainer Image(byte[] imageBytes, string fileName, Action<IImageDescriptor>? configure = null)
    {
        Guard.NotNull(imageBytes, nameof(imageBytes));
        if (imageBytes.Length == 0)
            throw new ArgumentException("Image byte array cannot be empty.", nameof(imageBytes));
        Guard.NotWhiteSpace(fileName, nameof(fileName));

        var image = new ImageElement
        {
            Bytes = imageBytes.ToArray(),
            FileName = fileName
        };
        configure?.Invoke(new ImageDescriptor(image, _theme));
        Element.Elements.Add(image);
        return this;
    }

    public IContainer Image(Stream imageStream, string fileName, Action<IImageDescriptor>? configure = null)
    {
        Guard.NotNull(imageStream, nameof(imageStream));
        if (!imageStream.CanRead)
            throw new ArgumentException("Image stream must be readable.", nameof(imageStream));
        Guard.NotWhiteSpace(fileName, nameof(fileName));

        using var ms = new MemoryStream();
        imageStream.CopyTo(ms);
        return Image(ms.ToArray(), fileName, configure);
    }

    public IContainer Line()
    {
        Element.Elements.Add(new LineElement());
        return this;
    }

    public IContainer PageBreak()
    {
        Element.Elements.Add(new PageBreakElement());
        return this;
    }

    public IContainer Component(IComponent component)
    {
        Guard.NotNull(component, nameof(component));
        component.Compose(this);
        return this;
    }
}

internal sealed class ImageDescriptor : IImageDescriptor
{
    private readonly ImageElement _image;
    private readonly DocumentTheme _theme;

    public ImageDescriptor(ImageElement image, DocumentTheme? theme = null)
    {
        _image = image;
        _theme = theme ?? DocumentTheme.Default;
    }

    public IImageDescriptor Width(float points)
    {
        _image.Width = Guard.PositiveFinite(points, nameof(points));
        return this;
    }

    public IImageDescriptor Height(float points)
    {
        _image.Height = Guard.PositiveFinite(points, nameof(points));
        return this;
    }

    public IImageDescriptor MaxWidth(float points)
    {
        _image.MaxWidth = Guard.PositiveFinite(points, nameof(points));
        return this;
    }

    public IImageDescriptor AltText(string text)
    {
        _image.AltText = Guard.NotNull(text, nameof(text));
        return this;
    }

    public IImageDescriptor Caption(string text, Action<ITextDescriptor>? configure = null)
    {
        Guard.NotNull(text, nameof(text));
        var descriptor = new TextDescriptor(text, _theme);
        descriptor.AlignCenter().FontSize(9).FontColor(Colors.Grey.L600).SpacingBefore(2).SpacingAfter(4);
        configure?.Invoke(descriptor);
        _image.Caption = descriptor.Element;
        return this;
    }

    public IImageDescriptor AlignLeft()
    {
        _image.Alignment = "left";
        return this;
    }

    public IImageDescriptor AlignCenter()
    {
        _image.Alignment = "center";
        return this;
    }

    public IImageDescriptor AlignRight()
    {
        _image.Alignment = "right";
        return this;
    }

    public IImageDescriptor WrapInline()
    {
        _image.WrapMode = "inline";
        return this;
    }

    public IImageDescriptor WrapSquare(float marginPoints = 6)
        => Wrap("square", marginPoints);

    public IImageDescriptor WrapTight(float marginPoints = 6)
        => Wrap("tight", marginPoints);

    public IImageDescriptor WrapTopBottom(float marginPoints = 6)
        => Wrap("topBottom", marginPoints);

    public IImageDescriptor BehindText()
    {
        _image.WrapMode = "behind";
        return this;
    }

    public IImageDescriptor InFrontOfText()
    {
        _image.WrapMode = "inFront";
        return this;
    }

    public IImageDescriptor FloatLeft(float marginPoints = 6)
        => Float("left", marginPoints);

    public IImageDescriptor FloatRight(float marginPoints = 6)
        => Float("right", marginPoints);

    public IImageDescriptor FloatCenter(float marginPoints = 6)
        => Float("center", marginPoints);

    public IImageDescriptor Position(float xPoints, float yPoints)
    {
        Guard.NonNegativeFinite(xPoints, nameof(xPoints));
        Guard.NonNegativeFinite(yPoints, nameof(yPoints));
        _image.WrapMode = _image.WrapMode == "inline" ? "square" : _image.WrapMode;
        _image.HorizontalRelativeFrom = "margin";
        _image.VerticalRelativeFrom = "paragraph";
        _image.HorizontalAlignment = null;
        _image.VerticalAlignment = null;
        _image.HorizontalOffset = xPoints;
        _image.VerticalOffset = yPoints;
        return this;
    }

    public IImageDescriptor PositionFromPage(float xPoints, float yPoints)
    {
        Position(xPoints, yPoints);
        _image.HorizontalRelativeFrom = "page";
        _image.VerticalRelativeFrom = "page";
        return this;
    }

    public IImageDescriptor Margin(float points)
        => Margin(points, points, points, points);

    public IImageDescriptor Margin(float topPoints, float rightPoints, float bottomPoints, float leftPoints)
    {
        _image.MarginTop = Guard.NonNegativeFinite(topPoints, nameof(topPoints));
        _image.MarginRight = Guard.NonNegativeFinite(rightPoints, nameof(rightPoints));
        _image.MarginBottom = Guard.NonNegativeFinite(bottomPoints, nameof(bottomPoints));
        _image.MarginLeft = Guard.NonNegativeFinite(leftPoints, nameof(leftPoints));
        return this;
    }

    public IImageDescriptor Border(float widthPoints = 1f, string hexColor = "000000")
    {
        _image.BorderWidth = Guard.NonNegativeFinite(widthPoints, nameof(widthPoints));
        _image.BorderColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public IImageDescriptor Rounded()
    {
        _image.Rounded = true;
        return this;
    }

    public IImageDescriptor Crop(float leftPercent, float topPercent, float rightPercent, float bottomPercent)
    {
        Guard.NonNegativeFinite(leftPercent, nameof(leftPercent));
        Guard.NonNegativeFinite(topPercent, nameof(topPercent));
        Guard.NonNegativeFinite(rightPercent, nameof(rightPercent));
        Guard.NonNegativeFinite(bottomPercent, nameof(bottomPercent));
        _image.CropLeftPercent = NormalizePercent(leftPercent);
        _image.CropTopPercent = NormalizePercent(topPercent);
        _image.CropRightPercent = NormalizePercent(rightPercent);
        _image.CropBottomPercent = NormalizePercent(bottomPercent);
        return this;
    }

    private IImageDescriptor Wrap(string mode, float marginPoints)
    {
        Guard.NonNegativeFinite(marginPoints, nameof(marginPoints));
        _image.WrapMode = mode;
        Margin(marginPoints);
        return this;
    }

    private IImageDescriptor Float(string alignment, float marginPoints)
    {
        Guard.NonNegativeFinite(marginPoints, nameof(marginPoints));
        _image.WrapMode = _image.WrapMode == "inline" ? "square" : _image.WrapMode;
        _image.HorizontalRelativeFrom = "margin";
        _image.VerticalRelativeFrom = "paragraph";
        _image.HorizontalAlignment = alignment;
        _image.VerticalAlignment = null;
        _image.HorizontalOffset = null;
        _image.VerticalOffset ??= 0;
        Margin(marginPoints);
        return this;
    }

    private static float NormalizePercent(float value)
        => MathCompat.Clamp(value, 0, 100);
}

internal sealed class ListDescriptor : IListDescriptor
{
    internal ListElement Element { get; }
    private readonly DocumentTheme _theme;

    public ListDescriptor(bool ordered, DocumentTheme? theme = null)
    {
        _theme = theme ?? DocumentTheme.Default;
        Element = new ListElement { Ordered = ordered };
    }

    public IListDescriptor Item(string text, Action<ITextDescriptor>? configure = null)
        => Item(text, 0, configure);

    public IListDescriptor Item(string text, int level, Action<ITextDescriptor>? configure = null)
    {
        var descriptor = new TextDescriptor(text, _theme);
        configure?.Invoke(descriptor);
        Element.Items.Add(new ListItemElement
        {
            Text = descriptor.Element,
            Level = NormalizeLevel(level)
        });
        return this;
    }

    public IListDescriptor Item(Action<ITextDescriptor> configure)
        => Item(0, configure);

    public IListDescriptor Item(int level, Action<ITextDescriptor> configure)
    {
        var descriptor = new TextDescriptor(string.Empty, _theme);
        configure(descriptor);
        Element.Items.Add(new ListItemElement
        {
            Text = descriptor.Element,
            Level = NormalizeLevel(level)
        });
        return this;
    }

    public IListDescriptor Marker(string marker, int level = 0, string? fontFamily = null)
    {
        Element.Markers[NormalizeLevel(level)] = new ListLevelMarker
        {
            Marker = marker,
            FontFamily = fontFamily
        };
        return this;
    }

    private static int NormalizeLevel(int level) => MathCompat.Clamp(level, 0, 8);
}

internal sealed class SeriesDescriptor : ISeriesDescriptor
{
    private readonly ChartSeriesElement _series;
    private string? _firstKind;

    public SeriesDescriptor(string name = "")
    {
        _series = new ChartSeriesElement { Name = name };
    }

    internal ChartSeriesElement Element => _series;

    private void SetKind(string kind)
    {
        if (_firstKind == null)
        {
            _firstKind = kind;
        }
        else if (_firstKind != kind)
        {
            throw new InvalidOperationException($"Series already uses '{_firstKind}' chart type; cannot mix chart types within a series. Start with {kind} points or use a new series.");
        }
        _series.Kind = kind;
    }

    ISeriesDescriptor ISeriesDescriptor.Bar(string label, double value)
    {
        SetKind("bar");
        _series.Points.Add(new ChartPointElement { Label = label, Value = value });
        return this;
    }

    ISeriesDescriptor ISeriesDescriptor.Line(string label, double value)
    {
        SetKind("line");
        _series.Points.Add(new ChartPointElement { Label = label, Value = value });
        return this;
    }

    ISeriesDescriptor ISeriesDescriptor.Pie(string label, double value)
    {
        SetKind("pie");
        _series.Points.Add(new ChartPointElement { Label = label, Value = value });
        return this;
    }

    ISeriesDescriptor ISeriesDescriptor.Doughnut(string label, double value)
    {
        SetKind("doughnut");
        _series.Points.Add(new ChartPointElement { Label = label, Value = value });
        return this;
    }

    ISeriesDescriptor ISeriesDescriptor.Color(string hexColor)
    {
        _series.Color = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }
}

internal sealed class ChartDescriptor : IChartDescriptor
{
    internal ChartElement Element { get; } = new();
    private string? _seriesKind;

    public ChartDescriptor(DocumentTheme? theme = null)
    {
        // Legacy support: initialize with default color (not used directly, but kept for compatibility)
    }

    IChartDescriptor IChartDescriptor.Title(string title)
    {
        Element.Title = title;
        return this;
    }

    IChartDescriptor IChartDescriptor.Series(Action<ISeriesDescriptor> configure)
    {
        var series = new SeriesDescriptor();
        configure(series);
        AddSeries(series.Element);
        return this;
    }

    IChartDescriptor IChartDescriptor.Series(string name, Action<ISeriesDescriptor> configure)
    {
        var series = new SeriesDescriptor(name);
        configure(series);
        AddSeries(series.Element);
        return this;
    }

    private void AddSeries(ChartSeriesElement series)
    {
        if (series.Points.Count > 0)
        {
            if (_seriesKind == null)
            {
                _seriesKind = series.Kind;
            }
            else if (_seriesKind != series.Kind)
            {
                throw new InvalidOperationException(
                    $"Chart already contains '{_seriesKind}' series; cannot add a '{series.Kind}' series. All series in a chart must use the same chart type.");
            }

            if ((_seriesKind == "pie" || _seriesKind == "doughnut") && Element.Series.Any(s => s.Points.Count > 0))
            {
                throw new InvalidOperationException("Pie and doughnut charts support only a single data series.");
            }
        }

        Element.Series.Add(series);
    }
}
