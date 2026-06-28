namespace TerraFluent.Docx.Reporting;

/// <summary>
/// A text highlight color, applied via <see cref="Infra.ITextDescriptor.Highlight"/>. These map to
/// Word's fixed highlight color palette (the same colors available from Word's highlighter tool),
/// not arbitrary hex colors.
/// </summary>
public enum HighlightColor
{
    /// <summary>Yellow highlight.</summary>
    Yellow,
    /// <summary>Green highlight.</summary>
    Green,
    /// <summary>Cyan highlight.</summary>
    Cyan,
    /// <summary>Magenta highlight.</summary>
    Magenta,
    /// <summary>Blue highlight.</summary>
    Blue,
    /// <summary>Red highlight.</summary>
    Red,
    /// <summary>Dark blue highlight.</summary>
    DarkBlue,
    /// <summary>Dark cyan highlight.</summary>
    DarkCyan,
    /// <summary>Dark green highlight.</summary>
    DarkGreen,
    /// <summary>Dark magenta highlight.</summary>
    DarkMagenta,
    /// <summary>Dark red highlight.</summary>
    DarkRed,
    /// <summary>Dark yellow highlight.</summary>
    DarkYellow,
    /// <summary>Dark gray highlight.</summary>
    DarkGray,
    /// <summary>Light gray highlight.</summary>
    LightGray,
    /// <summary>Black highlight.</summary>
    Black,
    /// <summary>White highlight.</summary>
    White,
    /// <summary>Removes highlighting.</summary>
    None
}

internal static class HighlightColorExtensions
{
    internal static string ToOoxmlValue(this HighlightColor color) => color switch
    {
        HighlightColor.Yellow => "yellow",
        HighlightColor.Green => "green",
        HighlightColor.Cyan => "cyan",
        HighlightColor.Magenta => "magenta",
        HighlightColor.Blue => "blue",
        HighlightColor.Red => "red",
        HighlightColor.DarkBlue => "darkBlue",
        HighlightColor.DarkCyan => "darkCyan",
        HighlightColor.DarkGreen => "darkGreen",
        HighlightColor.DarkMagenta => "darkMagenta",
        HighlightColor.DarkRed => "darkRed",
        HighlightColor.DarkYellow => "darkYellow",
        HighlightColor.DarkGray => "darkGray",
        HighlightColor.LightGray => "lightGray",
        HighlightColor.Black => "black",
        HighlightColor.White => "white",
        HighlightColor.None => "none",
        _ => "yellow"
    };
}
