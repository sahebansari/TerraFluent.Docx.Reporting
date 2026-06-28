using TerraFluent.Docx.Reporting.Core.Elements;
using TerraFluent.Docx.Reporting.Infra;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class TableDescriptor : ITableDescriptor
{
    private readonly DocumentTheme _theme;
    private readonly DocumentStyleCatalog _styles;
    internal TableElement Element { get; } = new();

    public TableDescriptor(DocumentTheme? theme = null, DocumentStyleCatalog? styles = null)
    {
        _theme = theme ?? DocumentTheme.Default;
        _styles = styles ?? new DocumentStyleCatalog();
        Element.BorderWidth = _theme.TableBorderWidth;
        Element.BorderColor = _theme.TableBorderColor;
        Element.CellPaddingTop = _theme.TableCellPaddingVertical;
        Element.CellPaddingRight = _theme.TableCellPaddingHorizontal;
        Element.CellPaddingBottom = _theme.TableCellPaddingVertical;
        Element.CellPaddingLeft = _theme.TableCellPaddingHorizontal;
        Element.HeaderBackgroundColor = _theme.TableHeaderBackgroundColor;
        Element.AlternateRowBackgroundColor = _theme.TableAlternateRowBackgroundColor;
        Element.RowMinHeight = _theme.TableRowMinHeight;
        Element.HeaderRowMinHeight = _theme.TableHeaderRowMinHeight;
    }

    public ITableDescriptor Style(string name)
    {
        Element.StyleId = DocumentStyleCatalog.StyleId(name);
        if (_styles.TryGetTableStyle(name, out var style))
            ApplyStyle(style.Style);
        return this;
    }

    public ITableDescriptor Width(float points)
    {
        Element.Width = points;
        Element.WidthType = "dxa";
        return this;
    }

    public ITableDescriptor WidthPercent(float percent)
    {
        Element.Width = MathCompat.Clamp(percent, 1, 100);
        Element.WidthType = "pct";
        return this;
    }

    public ITableDescriptor AlignLeft()
    {
        Element.Alignment = "left";
        return this;
    }

    public ITableDescriptor AlignCenter()
    {
        Element.Alignment = "center";
        return this;
    }

    public ITableDescriptor AlignRight()
    {
        Element.Alignment = "right";
        return this;
    }

    public ITableDescriptor ColumnsDefinition(Action<ITableColumnsDefinition> configure)
    {
        var def = new TableColumnsDefinitionDescriptor(Element);
        configure(def);
        return this;
    }

    public ITableDescriptor HeaderRow(Action<ITableRowDescriptor> configure)
    {
        var row = new TableRowDescriptor(_theme, _styles);
        configure(row);
        Element.HeaderRows.Add(row.Row);
        return this;
    }

    public ITableDescriptor Row(Action<ITableRowDescriptor> configure)
    {
        var row = new TableRowDescriptor(_theme, _styles);
        configure(row);
        Element.Rows.Add(row.Row);
        return this;
    }

    public ITableDescriptor Border(float width = 1f, string hexColor = "000000")
    {
        Element.BorderWidth = width;
        Element.BorderColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public ITableDescriptor CellPadding(float points)
    {
        Element.CellPaddingTop = points;
        Element.CellPaddingRight = points;
        Element.CellPaddingBottom = points;
        Element.CellPaddingLeft = points;
        return this;
    }

    public ITableDescriptor CellPadding(float verticalPoints, float horizontalPoints)
    {
        Element.CellPaddingTop = verticalPoints;
        Element.CellPaddingRight = horizontalPoints;
        Element.CellPaddingBottom = verticalPoints;
        Element.CellPaddingLeft = horizontalPoints;
        return this;
    }

    public ITableDescriptor HeaderBackground(string hexColor)
    {
        Element.HeaderBackgroundColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public ITableDescriptor AlternateRowBackground(string hexColor)
    {
        Element.AlternateRowBackgroundColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public ITableDescriptor RowMinHeight(float points)
    {
        Element.RowMinHeight = points;
        return this;
    }

    public ITableDescriptor HeaderRowMinHeight(float points)
    {
        Element.HeaderRowMinHeight = points;
        return this;
    }

    private void ApplyStyle(TableElement style)
    {
        Element.BorderWidth = style.BorderWidth;
        Element.BorderColor = style.BorderColor;
        Element.CellPaddingTop = style.CellPaddingTop;
        Element.CellPaddingRight = style.CellPaddingRight;
        Element.CellPaddingBottom = style.CellPaddingBottom;
        Element.CellPaddingLeft = style.CellPaddingLeft;
        Element.HeaderBackgroundColor = style.HeaderBackgroundColor;
        Element.AlternateRowBackgroundColor = style.AlternateRowBackgroundColor;
        Element.RowMinHeight = style.RowMinHeight;
        Element.HeaderRowMinHeight = style.HeaderRowMinHeight;
    }
}

internal sealed class TableColumnsDefinitionDescriptor : ITableColumnsDefinition
{
    private readonly TableElement _table;
    public TableColumnsDefinitionDescriptor(TableElement table) => _table = table;

    public ITableColumnsDefinition RelativeColumn(float size = 1)
    {
        _table.Columns.Add(new TableColumnDef { Mode = TableColumnDef.SizingMode.Relative, Size = size });
        return this;
    }

    public ITableColumnsDefinition ConstantColumn(float widthPoints)
    {
        _table.Columns.Add(new TableColumnDef { Mode = TableColumnDef.SizingMode.Constant, Size = widthPoints });
        return this;
    }
}

internal sealed class TableRowDescriptor : ITableRowDescriptor
{
    private readonly DocumentTheme _theme;
    private readonly DocumentStyleCatalog _styles;
    internal TableRow Row { get; } = new();

    public TableRowDescriptor(DocumentTheme? theme = null, DocumentStyleCatalog? styles = null)
    {
        _theme = theme ?? DocumentTheme.Default;
        _styles = styles ?? new DocumentStyleCatalog();
    }

    public ITableRowDescriptor KeepTogether(bool value = true)
    {
        Row.KeepTogether = value;
        return this;
    }

    public ITableCellDescriptor Cell(int columnSpan = 1)
    {
        var container = new ContainerDescriptor(_theme, _styles);
        var cell = new TableCell
        {
            Container = container.Element,
            ColumnSpan = Math.Max(1, columnSpan)
        };
        Row.Cells.Add(cell);
        return new TableCellDescriptor(cell, container);
    }
}

internal sealed class TableCellDescriptor : ITableCellDescriptor
{
    private readonly TableCell _cell;
    private readonly ContainerDescriptor _container;

    public TableCellDescriptor(TableCell cell, ContainerDescriptor container)
    {
        _cell = cell;
        _container = container;
    }

    public ITableCellDescriptor ColumnSpan(int columns)
    {
        _cell.ColumnSpan = Math.Max(1, columns);
        return this;
    }

    public ITableCellDescriptor VerticalMergeStart()
    {
        _cell.VerticalMerge = "restart";
        return this;
    }

    public ITableCellDescriptor VerticalMergeContinue()
    {
        _cell.VerticalMerge = "continue";
        return this;
    }

    public ITableCellDescriptor Background(string hexColor)
    {
        _cell.BackgroundColor = HexColor.Validate(hexColor, nameof(hexColor));
        return this;
    }

    public ITableCellDescriptor Border(float width = 1f, string hexColor = "000000")
    {
        BorderTop(width, hexColor);
        BorderRight(width, hexColor);
        BorderBottom(width, hexColor);
        BorderLeft(width, hexColor);
        return this;
    }

    public ITableCellDescriptor BorderTop(float width = 1f, string hexColor = "000000")
    {
        _cell.TopBorder = new CellBorder { Width = width, Color = HexColor.Validate(hexColor, nameof(hexColor)) };
        return this;
    }

    public ITableCellDescriptor BorderRight(float width = 1f, string hexColor = "000000")
    {
        _cell.RightBorder = new CellBorder { Width = width, Color = HexColor.Validate(hexColor, nameof(hexColor)) };
        return this;
    }

    public ITableCellDescriptor BorderBottom(float width = 1f, string hexColor = "000000")
    {
        _cell.BottomBorder = new CellBorder { Width = width, Color = HexColor.Validate(hexColor, nameof(hexColor)) };
        return this;
    }

    public ITableCellDescriptor BorderLeft(float width = 1f, string hexColor = "000000")
    {
        _cell.LeftBorder = new CellBorder { Width = width, Color = HexColor.Validate(hexColor, nameof(hexColor)) };
        return this;
    }

    public ITableCellDescriptor Padding(float points)
    {
        _cell.PaddingTop = points;
        _cell.PaddingRight = points;
        _cell.PaddingBottom = points;
        _cell.PaddingLeft = points;
        return this;
    }

    public ITableCellDescriptor Padding(float verticalPoints, float horizontalPoints)
    {
        _cell.PaddingTop = verticalPoints;
        _cell.PaddingRight = horizontalPoints;
        _cell.PaddingBottom = verticalPoints;
        _cell.PaddingLeft = horizontalPoints;
        return this;
    }

    public ITableCellDescriptor VerticalAlignTop()
    {
        _cell.VerticalAlignment = "top";
        return this;
    }

    public ITableCellDescriptor VerticalAlignMiddle()
    {
        _cell.VerticalAlignment = "center";
        return this;
    }

    public ITableCellDescriptor VerticalAlignBottom()
    {
        _cell.VerticalAlignment = "bottom";
        return this;
    }

    public ITableCellDescriptor TextDirectionLeftToRight()
    {
        _cell.TextDirection = "lrTb";
        return this;
    }

    public ITableCellDescriptor TextDirectionTopToBottom()
    {
        _cell.TextDirection = "tbRl";
        return this;
    }

    public ITableCellDescriptor TextDirectionBottomToTop()
    {
        _cell.TextDirection = "btLr";
        return this;
    }

    public IContainer H1(string text) => _container.H1(text);
    public IContainer H2(string text) => _container.H2(text);
    public IContainer H3(string text) => _container.H3(text);
    public IContainer H4(string text) => _container.H4(text);
    public IContainer H5(string text) => _container.H5(text);
    public IContainer H6(string text) => _container.H6(text);
    public IContainer Column(Action<IColumnDescriptor> configure) => _container.Column(configure);
    public IContainer Row(Action<IRowDescriptor> configure) => _container.Row(configure);
    public IContainer Text(string text, Action<ITextDescriptor>? configure = null) => _container.Text(text, configure);
    public IContainer Text(Action<ITextDescriptor> configure) => _container.Text(configure);
    public IContainer Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null) => _container.Hyperlink(text, url, configure);
    public IContainer Bookmark(string name) => _container.Bookmark(name);
    public IContainer Bookmark(string name, string text, Action<ITextDescriptor>? configure = null) => _container.Bookmark(name, text, configure);
    public IContainer TableOfContents(string title = "Contents", int minLevel = 1, int maxLevel = 3) => _container.TableOfContents(title, minLevel, maxLevel);
    public IContainer BulletList(Action<IListDescriptor> configure) => _container.BulletList(configure);
    public IContainer NumberedList(Action<IListDescriptor> configure) => _container.NumberedList(configure);
    public IContainer Table(Action<ITableDescriptor> configure) => _container.Table(configure);
    public IContainer Chart(Action<IChartDescriptor> configure) => _container.Chart(configure);
    public IContainer Image(string filePath, float? width = null) => _container.Image(filePath, width);
    public IContainer Image(string filePath, Action<IImageDescriptor> configure) => _container.Image(filePath, configure);
    public IContainer Image(byte[] imageBytes, string fileName, Action<IImageDescriptor>? configure = null) => _container.Image(imageBytes, fileName, configure);
    public IContainer Image(Stream imageStream, string fileName, Action<IImageDescriptor>? configure = null) => _container.Image(imageStream, fileName, configure);
    public IContainer Line() => _container.Line();
    public IContainer PageBreak() => _container.PageBreak();
    public IContainer Component(IComponent component) => _container.Component(component);
}
