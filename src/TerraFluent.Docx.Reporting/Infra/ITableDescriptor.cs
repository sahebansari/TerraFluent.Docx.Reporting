namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Builds a table inside an <see cref="IContainer.Table"/> call: columns, header and body rows,
/// and table-wide styling.
/// </summary>
public interface ITableDescriptor
{
    /// <summary>Applies a registered named table style. The style must have been registered via
    /// <see cref="IDocumentContainer.TableStyle"/>.</summary>
    ITableDescriptor Style(string name);

    /// <summary>Sets a fixed table width in points.</summary>
    ITableDescriptor Width(float points);

    /// <summary>Sets the table width as a percentage (1-100) of the available content width.</summary>
    ITableDescriptor WidthPercent(float percent);

    /// <summary>Left-aligns the table on the page.</summary>
    ITableDescriptor AlignLeft();

    /// <summary>Center-aligns the table on the page.</summary>
    ITableDescriptor AlignCenter();

    /// <summary>Right-aligns the table on the page.</summary>
    ITableDescriptor AlignRight();

    /// <summary>
    /// Defines the table's columns as relative or fixed widths. If never called, column widths are
    /// inferred automatically from the widest header or body row when the document is published.
    /// </summary>
    /// <param name="configure">Adds columns in left-to-right order. Cannot be null.</param>
    ITableDescriptor ColumnsDefinition(Action<ITableColumnsDefinition> configure);

    /// <summary>Adds a repeating header row, shown again at the top of each page the table spans.</summary>
    /// <param name="configure">Adds cells to the row. Cannot be null.</param>
    ITableDescriptor HeaderRow(Action<ITableRowDescriptor> configure);

    /// <summary>Adds a body row.</summary>
    /// <param name="configure">Adds cells to the row. Cannot be null.</param>
    ITableDescriptor Row(Action<ITableRowDescriptor> configure);

    /// <summary>Sets the table's default cell border on all sides.</summary>
    /// <param name="width">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableDescriptor Border(float width = 1f, string hexColor = "000000");

    /// <summary>Sets uniform default cell padding, in points, on all sides.</summary>
    ITableDescriptor CellPadding(float points);

    /// <summary>Sets default cell padding, in points, separately for top/bottom and left/right.</summary>
    ITableDescriptor CellPadding(float verticalPoints, float horizontalPoints);

    /// <summary>Sets the background fill color for header rows.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableDescriptor HeaderBackground(string hexColor);

    /// <summary>Sets the background fill color applied to alternating body rows (banded rows).</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableDescriptor AlternateRowBackground(string hexColor);

    /// <summary>Sets the minimum height, in points, for body rows.</summary>
    ITableDescriptor RowMinHeight(float points);

    /// <summary>Sets the minimum height, in points, for header rows.</summary>
    ITableDescriptor HeaderRowMinHeight(float points);
}

/// <summary>
/// Defines a table's columns, in left-to-right order, inside an <see cref="ITableDescriptor.ColumnsDefinition"/> call.
/// A table's columns must be either all relative, all constant, or a mix of both; mixing computes each
/// column's share of the table's total width (its explicit <see cref="ITableDescriptor.Width(float)"/> or the page content width).
/// </summary>
public interface ITableColumnsDefinition
{
    /// <summary>Adds a column whose width is proportional to other relative columns in the table.</summary>
    /// <param name="size">The column's relative size weight compared to other <see cref="RelativeColumn"/> calls in the same table.</param>
    ITableColumnsDefinition RelativeColumn(float size = 1);

    /// <summary>Adds a column with a fixed width.</summary>
    /// <param name="widthPoints">The column's width in points.</param>
    ITableColumnsDefinition ConstantColumn(float widthPoints);
}

/// <summary>
/// Builds a single table row inside an <see cref="ITableDescriptor.Row"/> or
/// <see cref="ITableDescriptor.HeaderRow"/> call.
/// </summary>
public interface ITableRowDescriptor
{
    /// <summary>Prevents this row from splitting across a page break.</summary>
    ITableRowDescriptor KeepTogether(bool value = true);

    /// <summary>Adds a cell to the row and returns its container for adding content.</summary>
    /// <param name="columnSpan">The number of columns this cell spans, starting at the next unfilled column.</param>
    ITableCellDescriptor Cell(int columnSpan = 1);
}

/// <summary>
/// Configures a single table cell's layout and styling. Inherits <see cref="IContainer"/>, so cell
/// content is added with the same content methods used elsewhere in the document.
/// </summary>
public interface ITableCellDescriptor : IContainer
{
    /// <summary>Sets the number of columns this cell spans.</summary>
    ITableCellDescriptor ColumnSpan(int columns);

    /// <summary>Marks this cell as the start of a vertical merge; cells below it use <see cref="VerticalMergeContinue"/>
    /// to merge into it. The merged cells must share the same column position.</summary>
    ITableCellDescriptor VerticalMergeStart();

    /// <summary>Merges this cell into the vertical merge started by the cell directly above it in the same column.</summary>
    ITableCellDescriptor VerticalMergeContinue();

    /// <summary>Sets the cell's background fill color.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableCellDescriptor Background(string hexColor);

    /// <summary>Sets a border on all four sides of the cell.</summary>
    /// <param name="width">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableCellDescriptor Border(float width = 1f, string hexColor = "000000");

    /// <summary>Sets a border on the top of the cell.</summary>
    /// <param name="width">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableCellDescriptor BorderTop(float width = 1f, string hexColor = "000000");

    /// <summary>Sets a border on the right of the cell.</summary>
    /// <param name="width">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableCellDescriptor BorderRight(float width = 1f, string hexColor = "000000");

    /// <summary>Sets a border on the bottom of the cell.</summary>
    /// <param name="width">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableCellDescriptor BorderBottom(float width = 1f, string hexColor = "000000");

    /// <summary>Sets a border on the left of the cell.</summary>
    /// <param name="width">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    ITableCellDescriptor BorderLeft(float width = 1f, string hexColor = "000000");

    /// <summary>Sets uniform cell padding, in points, on all sides, overriding the table default.</summary>
    ITableCellDescriptor Padding(float points);

    /// <summary>Sets cell padding, in points, separately for top/bottom and left/right, overriding the table default.</summary>
    ITableCellDescriptor Padding(float verticalPoints, float horizontalPoints);

    /// <summary>Aligns cell content to the top of the cell.</summary>
    ITableCellDescriptor VerticalAlignTop();

    /// <summary>Aligns cell content to the vertical middle of the cell.</summary>
    ITableCellDescriptor VerticalAlignMiddle();

    /// <summary>Aligns cell content to the bottom of the cell.</summary>
    ITableCellDescriptor VerticalAlignBottom();

    /// <summary>Sets normal left-to-right, top-to-bottom text direction for the cell.</summary>
    ITableCellDescriptor TextDirectionLeftToRight();

    /// <summary>Rotates cell text to read top-to-bottom (90 degrees clockwise).</summary>
    ITableCellDescriptor TextDirectionTopToBottom();

    /// <summary>Rotates cell text to read bottom-to-top (90 degrees counter-clockwise).</summary>
    ITableCellDescriptor TextDirectionBottomToTop();
}
