namespace TerraFluent.Docx.Reporting.Core;

internal static class Guard
{
    public static T NotNull<T>(T? value, string paramName) where T : class
    {
        if (value == null)
            throw new ArgumentNullException(paramName);

        return value!;
    }

    public static string NotWhiteSpace(string? value, string paramName)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Value cannot be null, empty, or white space.", paramName);

        return value!;
    }

    public static float PositiveFinite(float value, string paramName)
    {
        if (float.IsNaN(value) || float.IsInfinity(value) || value <= 0)
            throw new ArgumentOutOfRangeException(paramName, value, "Value must be a positive finite number.");

        return value;
    }

    public static float NonNegativeFinite(float value, string paramName)
    {
        if (float.IsNaN(value) || float.IsInfinity(value) || value < 0)
            throw new ArgumentOutOfRangeException(paramName, value, "Value must be a non-negative finite number.");

        return value;
    }

    public static int InRange(int value, int min, int max, string paramName)
    {
        if (value < min || value > max)
            throw new ArgumentOutOfRangeException(paramName, value, $"Value must be between {min} and {max}.");

        return value;
    }
}
