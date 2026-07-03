namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Configures a barcode added via one of the <see cref="IContainer"/> <c>Barcode</c> overloads:
/// sizing, alignment, bar color, the human-readable text line, and captions.
/// </summary>
public interface IBarcodeDescriptor
{
    /// <summary>Sets the rendered width of the barcode (bars plus the space between them), in points.</summary>
    IBarcodeDescriptor Width(float points);

    /// <summary>Sets the rendered height of the bars, in points.</summary>
    IBarcodeDescriptor Height(float points);

    /// <summary>Caps the rendered width, in points, scaling the bars down if they would otherwise be wider.</summary>
    IBarcodeDescriptor MaxWidth(float points);

    /// <summary>Sets accessibility (alt) text for the barcode.</summary>
    IBarcodeDescriptor AltText(string text);

    /// <summary>Shows or hides the human-readable value printed below the bars. Shown by default.</summary>
    /// <param name="show">Whether to show the human-readable text line.</param>
    IBarcodeDescriptor ShowText(bool show = true);

    /// <summary>Sets the bar color.</summary>
    /// <param name="hexColor">A six-digit hex color (with or without a leading "#") or three-digit CSS shorthand.</param>
    IBarcodeDescriptor BarColor(string hexColor);

    /// <summary>Left-aligns the barcode's paragraph.</summary>
    IBarcodeDescriptor AlignLeft();

    /// <summary>Center-aligns the barcode's paragraph.</summary>
    IBarcodeDescriptor AlignCenter();

    /// <summary>Right-aligns the barcode's paragraph.</summary>
    IBarcodeDescriptor AlignRight();

    /// <summary>Adds a caption paragraph below the barcode (and its human-readable text line, if shown).</summary>
    /// <param name="text">The caption text.</param>
    /// <param name="configure">Optional formatting for the caption paragraph, applied after the default caption style.</param>
    IBarcodeDescriptor Caption(string text, Action<ITextDescriptor>? configure = null);
}
