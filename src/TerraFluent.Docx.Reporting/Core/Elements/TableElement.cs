namespace TerraFluent.Docx.Reporting.Core.Elements;

internal sealed class TableElement : IElement
{
    public string? StyleId { get; set; }
    public List<TableColumnDef> Columns { get; } = [];
    public List<TableRow> HeaderRows { get; } = [];
    public List<TableRow> Rows { get; } = [];
    public float BorderWidth { get; set; }
    public string BorderColor { get; set; } = "000000";
    public float CellPaddingTop { get; set; } = 3;
    public float CellPaddingRight { get; set; } = 5;
    public float CellPaddingBottom { get; set; } = 3;
    public float CellPaddingLeft { get; set; } = 5;
    public string? HeaderBackgroundColor { get; set; }
    public string? AlternateRowBackgroundColor { get; set; }
    public float RowMinHeight { get; set; }
    public float HeaderRowMinHeight { get; set; }
    public float? Width { get; set; }
    public string WidthType { get; set; } = "pct";
    public string Alignment { get; set; } = "left";
    public string? CaptionLabel { get; set; }  // "Table" for auto-numbered
    public string? CaptionDescription { get; set; }  // Description text after number
    public int? CaptionTableNumber { get; set; }  // Auto-assigned table number
}

internal sealed class TableColumnDef
{
    public enum SizingMode { Relative, Constant }
    public SizingMode Mode { get; set; } = SizingMode.Relative;
    public float Size { get; set; } = 1;
}

internal sealed class TableRow
{
    public List<TableCell> Cells { get; } = [];
    public bool KeepTogether { get; set; }
}

internal sealed class TableCell
{
    public ContainerElement Container { get; set; } = new();
    public int ColumnSpan { get; set; } = 1;
    public string? BackgroundColor { get; set; }
    public float? PaddingTop { get; set; }
    public float? PaddingRight { get; set; }
    public float? PaddingBottom { get; set; }
    public float? PaddingLeft { get; set; }
    public string? VerticalAlignment { get; set; }
    public string? VerticalMerge { get; set; }
    public string? TextDirection { get; set; }
    public CellBorder? TopBorder { get; set; }
    public CellBorder? RightBorder { get; set; }
    public CellBorder? BottomBorder { get; set; }
    public CellBorder? LeftBorder { get; set; }
}

internal sealed class CellBorder
{
    public float Width { get; set; }
    public string Color { get; set; } = "000000";
}
