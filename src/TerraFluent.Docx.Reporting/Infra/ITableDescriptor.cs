namespace TerraFluent.Docx.Reporting.Infra;

public interface ITableDescriptor
{
    ITableDescriptor Style(string name);
    ITableDescriptor Width(float points);
    ITableDescriptor WidthPercent(float percent);
    ITableDescriptor AlignLeft();
    ITableDescriptor AlignCenter();
    ITableDescriptor AlignRight();
    ITableDescriptor ColumnsDefinition(Action<ITableColumnsDefinition> configure);
    ITableDescriptor HeaderRow(Action<ITableRowDescriptor> configure);
    ITableDescriptor Row(Action<ITableRowDescriptor> configure);
    ITableDescriptor Border(float width = 1f, string hexColor = "000000");
    ITableDescriptor CellPadding(float points);
    ITableDescriptor CellPadding(float verticalPoints, float horizontalPoints);
    ITableDescriptor HeaderBackground(string hexColor);
    ITableDescriptor AlternateRowBackground(string hexColor);
    ITableDescriptor RowMinHeight(float points);
    ITableDescriptor HeaderRowMinHeight(float points);
}

public interface ITableColumnsDefinition
{
    ITableColumnsDefinition RelativeColumn(float size = 1);
    ITableColumnsDefinition ConstantColumn(float widthPoints);
}

public interface ITableRowDescriptor
{
    ITableRowDescriptor KeepTogether(bool value = true);
    ITableCellDescriptor Cell(int columnSpan = 1);
}

public interface ITableCellDescriptor : IContainer
{
    ITableCellDescriptor ColumnSpan(int columns);
    ITableCellDescriptor VerticalMergeStart();
    ITableCellDescriptor VerticalMergeContinue();
    ITableCellDescriptor Background(string hexColor);
    ITableCellDescriptor Border(float width = 1f, string hexColor = "000000");
    ITableCellDescriptor BorderTop(float width = 1f, string hexColor = "000000");
    ITableCellDescriptor BorderRight(float width = 1f, string hexColor = "000000");
    ITableCellDescriptor BorderBottom(float width = 1f, string hexColor = "000000");
    ITableCellDescriptor BorderLeft(float width = 1f, string hexColor = "000000");
    ITableCellDescriptor Padding(float points);
    ITableCellDescriptor Padding(float verticalPoints, float horizontalPoints);
    ITableCellDescriptor VerticalAlignTop();
    ITableCellDescriptor VerticalAlignMiddle();
    ITableCellDescriptor VerticalAlignBottom();
    ITableCellDescriptor TextDirectionLeftToRight();
    ITableCellDescriptor TextDirectionTopToBottom();
    ITableCellDescriptor TextDirectionBottomToTop();
}
