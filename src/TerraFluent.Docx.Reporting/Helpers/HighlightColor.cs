namespace TerraFluent.Docx.Reporting;

public enum HighlightColor
{
    Yellow,
    Green,
    Cyan,
    Magenta,
    Blue,
    Red,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
    Black,
    White,
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
