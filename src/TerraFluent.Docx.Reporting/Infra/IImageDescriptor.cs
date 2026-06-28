namespace TerraFluent.Docx.Reporting.Infra;

public interface IImageDescriptor
{
    IImageDescriptor Width(float points);
    IImageDescriptor Height(float points);
    IImageDescriptor MaxWidth(float points);
    IImageDescriptor AltText(string text);
    IImageDescriptor Caption(string text, Action<ITextDescriptor>? configure = null);
    IImageDescriptor AlignLeft();
    IImageDescriptor AlignCenter();
    IImageDescriptor AlignRight();
    IImageDescriptor WrapInline();
    IImageDescriptor WrapSquare(float marginPoints = 6);
    IImageDescriptor WrapTight(float marginPoints = 6);
    IImageDescriptor WrapTopBottom(float marginPoints = 6);
    IImageDescriptor BehindText();
    IImageDescriptor InFrontOfText();
    IImageDescriptor FloatLeft(float marginPoints = 6);
    IImageDescriptor FloatRight(float marginPoints = 6);
    IImageDescriptor FloatCenter(float marginPoints = 6);
    IImageDescriptor Position(float xPoints, float yPoints);
    IImageDescriptor PositionFromPage(float xPoints, float yPoints);
    IImageDescriptor Margin(float points);
    IImageDescriptor Margin(float topPoints, float rightPoints, float bottomPoints, float leftPoints);
    IImageDescriptor Border(float widthPoints = 1f, string hexColor = "000000");
    IImageDescriptor Rounded();
    IImageDescriptor Crop(float leftPercent, float topPercent, float rightPercent, float bottomPercent);
}
