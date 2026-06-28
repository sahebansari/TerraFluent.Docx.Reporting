namespace TerraFluent.Docx.Reporting.Infra;

public interface IPageDescriptor
{
    IPageDescriptor Size(PageSize size);
    IPageDescriptor Size(float widthPoints, float heightPoints);
    IPageDescriptor Landscape();
    IPageDescriptor Portrait();
    IPageDescriptor Margin(float allPoints);
    IPageDescriptor Margin(float verticalPoints, float horizontalPoints);
    IPageDescriptor Margin(float topPoints, float rightPoints, float bottomPoints, float leftPoints);
    IPageDescriptor MarginTop(float points);
    IPageDescriptor MarginRight(float points);
    IPageDescriptor MarginBottom(float points);
    IPageDescriptor MarginLeft(float points);
    IPageDescriptor DefaultTextStyle(Action<ITextDescriptor> configure);
    IPageDescriptor PageNumberStart(int value);
    /// <summary>
    /// Sets the page number format for this section to an OOXML ST_NumberFormat value
    /// (e.g., "lowerRoman" for i, ii, iii). The value is written as-is to the
    /// <c>w:fmt</c> attribute, so it must be one of the values defined by that
    /// enumeration or Word will treat the document as invalid.
    /// Common formats: "decimal" (default), "lowerRoman", "upperRoman", "lowerLetter", "upperLetter", "none".
    /// </summary>
    IPageDescriptor PageNumberFormat(string format);
    /// <summary>
    /// Sets the document's page background color (Word's "Page Color"). This maps to the
    /// single, document-wide &lt;w:background&gt; element in OOXML - there is no per-section
    /// or per-page equivalent, so the color cannot vary across pages within one .docx.
    /// If multiple pages call this, only the first non-empty value (in page order) is used
    /// and later calls are ignored.
    /// </summary>
    IPageDescriptor Background(string hexColor);
    IPageDescriptor Watermark(string text, string hexColor = "D9D9D9", float fontSize = 54);
    IPageDescriptor Columns(int count, float spacingPoints = 36, bool separatorLine = false);
    IPageDescriptor SingleColumn();
    IContainer Header();
    IContainer OddPageHeader();
    IContainer FirstPageHeader();
    IContainer EvenPageHeader();
    IContainer Content();
    IContainer Footer();
    IContainer OddPageFooter();
    IContainer FirstPageFooter();
    IContainer EvenPageFooter();
}
