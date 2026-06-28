namespace TerraFluent.Docx.Reporting.Infra;

public interface IDocumentThemeDescriptor
{
    IDocumentThemeDescriptor DefaultFont(string family, float size = 11);
    IDocumentThemeDescriptor DefaultFontFamily(string family);
    IDocumentThemeDescriptor DefaultFontSize(float size);
    IDocumentThemeDescriptor DefaultTextColor(string hexColor);
    IDocumentThemeDescriptor HeadingColor(string hexColor);
    IDocumentThemeDescriptor AccentColor(string hexColor);
    IDocumentThemeDescriptor HyperlinkColor(string hexColor);
    IDocumentThemeDescriptor TableHeaderBackground(string hexColor);
    IDocumentThemeDescriptor TableAlternateRowBackground(string hexColor);
    IDocumentThemeDescriptor TableBorder(float width, string hexColor);
    IDocumentThemeDescriptor TableCellPadding(float points);
    IDocumentThemeDescriptor TableCellPadding(float verticalPoints, float horizontalPoints);
    IDocumentThemeDescriptor TableRowMinHeight(float points);
    IDocumentThemeDescriptor TableHeaderRowMinHeight(float points);
}
