namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Configures an image added via one of the <see cref="IContainer"/> <c>Image</c> overloads: sizing,
/// alignment, text wrapping, floating position, borders, rounding, and cropping.
/// </summary>
public interface IImageDescriptor
{
    /// <summary>Sets the rendered width, in points. The aspect ratio is preserved unless <see cref="Height"/> is also set.</summary>
    IImageDescriptor Width(float points);

    /// <summary>Sets the rendered height, in points. The aspect ratio is preserved unless <see cref="Width"/> is also set.</summary>
    IImageDescriptor Height(float points);

    /// <summary>Caps the rendered width, in points, scaling the image down (preserving aspect ratio) if it would otherwise be wider.</summary>
    IImageDescriptor MaxWidth(float points);

    /// <summary>Sets accessibility (alt) text for the image.</summary>
    IImageDescriptor AltText(string text);

    /// <summary>Adds a caption paragraph below the image.</summary>
    /// <param name="text">The caption text.</param>
    /// <param name="configure">Optional formatting for the caption paragraph, applied after the default caption style.</param>
    IImageDescriptor Caption(string text, Action<ITextDescriptor>? configure = null);

    /// <summary>
    /// Left-aligns the image. For an inline image, aligns its paragraph. For a floating image with no
    /// explicit <see cref="Position"/>/<see cref="PositionFromPage"/>/float call, sets its horizontal position.
    /// </summary>
    IImageDescriptor AlignLeft();

    /// <summary>Center-aligns the image (see <see cref="AlignLeft"/> for inline vs. floating behavior).</summary>
    IImageDescriptor AlignCenter();

    /// <summary>Right-aligns the image (see <see cref="AlignLeft"/> for inline vs. floating behavior).</summary>
    IImageDescriptor AlignRight();

    /// <summary>Renders the image inline with surrounding text, as if it were a character in the paragraph.</summary>
    IImageDescriptor WrapInline();

    /// <summary>Makes the image float, with surrounding text wrapping around its bounding box on all sides.</summary>
    /// <param name="marginPoints">Spacing, in points, kept clear between the image and the wrapping text.</param>
    IImageDescriptor WrapSquare(float marginPoints = 6);

    /// <summary>Makes the image float, with surrounding text wrapping tightly around its visible edges.</summary>
    /// <param name="marginPoints">Spacing, in points, kept clear between the image and the wrapping text.</param>
    IImageDescriptor WrapTight(float marginPoints = 6);

    /// <summary>Makes the image float, with surrounding text only above and below it (no wrap on the sides).</summary>
    /// <param name="marginPoints">Spacing, in points, kept clear above and below the image.</param>
    IImageDescriptor WrapTopBottom(float marginPoints = 6);

    /// <summary>Makes the image float behind the document text, as a background layer.</summary>
    IImageDescriptor BehindText();

    /// <summary>Makes the image float in front of the document text, on top of it.</summary>
    IImageDescriptor InFrontOfText();

    /// <summary>Floats the image at the left margin.</summary>
    /// <param name="marginPoints">Spacing, in points, kept clear between the image and surrounding text.</param>
    IImageDescriptor FloatLeft(float marginPoints = 6);

    /// <summary>Floats the image at the right margin.</summary>
    /// <param name="marginPoints">Spacing, in points, kept clear between the image and surrounding text.</param>
    IImageDescriptor FloatRight(float marginPoints = 6);

    /// <summary>Floats the image horizontally centered between the margins.</summary>
    /// <param name="marginPoints">Spacing, in points, kept clear between the image and surrounding text.</param>
    IImageDescriptor FloatCenter(float marginPoints = 6);

    /// <summary>Positions a floating image at an exact offset from the page margin and the current paragraph.</summary>
    /// <param name="xPoints">Horizontal offset, in points, from the margin.</param>
    /// <param name="yPoints">Vertical offset, in points, from the current paragraph position.</param>
    IImageDescriptor Position(float xPoints, float yPoints);

    /// <summary>Positions a floating image at an exact offset from the page edge, ignoring margins and paragraph flow.</summary>
    /// <param name="xPoints">Horizontal offset, in points, from the left edge of the page.</param>
    /// <param name="yPoints">Vertical offset, in points, from the top edge of the page.</param>
    IImageDescriptor PositionFromPage(float xPoints, float yPoints);

    /// <summary>Sets uniform wrap margin, in points, on all sides of a floating image.</summary>
    IImageDescriptor Margin(float points);

    /// <summary>Sets wrap margin, in points, independently on each side of a floating image.</summary>
    IImageDescriptor Margin(float topPoints, float rightPoints, float bottomPoints, float leftPoints);

    /// <summary>Adds a border around the image.</summary>
    /// <param name="widthPoints">Border line width, in points.</param>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IImageDescriptor Border(float widthPoints = 1f, string hexColor = "000000");

    /// <summary>Renders the image with rounded corners.</summary>
    IImageDescriptor Rounded();

    /// <summary>Crops the image by the given percentage from each edge.</summary>
    /// <param name="leftPercent">Percentage (0-100) to crop from the left edge.</param>
    /// <param name="topPercent">Percentage (0-100) to crop from the top edge.</param>
    /// <param name="rightPercent">Percentage (0-100) to crop from the right edge.</param>
    /// <param name="bottomPercent">Percentage (0-100) to crop from the bottom edge.</param>
    IImageDescriptor Crop(float leftPercent, float topPercent, float rightPercent, float bottomPercent);
}
