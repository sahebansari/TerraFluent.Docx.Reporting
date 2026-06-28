namespace TerraFluent.Docx.Reporting;

// Math.Clamp isn't part of netstandard2.0; this mirrors its semantics for every TFM the library targets.
internal static class MathCompat
{
    public static T Clamp<T>(T value, T min, T max) where T : IComparable<T>
    {
        if (min.CompareTo(max) > 0)
            throw new ArgumentException($"'{min}' cannot be greater than '{max}'.", nameof(min));

        if (value.CompareTo(min) < 0) return min;
        if (value.CompareTo(max) > 0) return max;
        return value;
    }
}
