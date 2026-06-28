namespace TerraFluent.Docx.Reporting;

internal static class HexColor
{
    public static string Validate(string hexColor, string paramName)
    {
        if (string.IsNullOrWhiteSpace(hexColor))
            throw new ArgumentException("Hex color cannot be null or empty.", paramName);

        var value = hexColor.TrimStart('#');

        if (value.Length == 3 && IsHex(value))
            value = $"{value[0]}{value[0]}{value[1]}{value[1]}{value[2]}{value[2]}";

        if (value.Length != 6 || !IsHex(value))
            throw new ArgumentException(
                $"'{hexColor}' is not a valid hex color (expected 6 hex digits, " +
                $"e.g. \"1976D2\", or 3-digit shorthand, e.g. \"D32\").",
                paramName);

        return value.ToUpperInvariant();
    }

    private static bool IsHex(string value) => value.All(Uri.IsHexDigit);
}
