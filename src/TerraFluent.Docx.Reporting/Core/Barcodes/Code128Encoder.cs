namespace TerraFluent.Docx.Reporting.Core.Barcodes;

/// <summary>
/// Encodes text as Code 128 (Subset B) bar/space widths, per ISO/IEC 15417.
/// </summary>
internal static class Code128Encoder
{
    private const int StartB = 104;
    private const int Stop = 106;

    // ISO/IEC 15417 Code 128 symbol character bar/space width table, indexed by symbol value
    // (0-106). Each entry is a sequence of module widths alternating bar, space, bar, space, ...
    // starting with a bar. All entries are 6 widths (11 modules) except the STOP pattern at
    // index 106, which is 7 widths (13 modules, the last of which is the trailing stop bar).
    private static readonly string[] Patterns =
    [
        "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213",
        "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132",
        "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211",
        "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
        "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331",
        "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111",
        "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214",
        "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
        "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141",
        "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141",
        "114131", "311141", "411131", "211412", "211214", "211232", "2331112",
    ];

    /// <summary>
    /// Encodes <paramref name="value"/> as Code 128 Subset B and returns the bar/space module
    /// widths, in modules, starting with a bar and alternating bar/space thereafter.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> is null, empty, or contains a character outside ASCII 32-126.
    /// </exception>
    public static int[] Encode(string value)
    {
        if (string.IsNullOrEmpty(value))
            throw new ArgumentException("Barcode value cannot be null or empty.", nameof(value));

        var values = new int[value.Length];
        for (int i = 0; i < value.Length; i++)
        {
            int code = value[i];
            if (code is < 32 or > 126)
                throw new ArgumentException(
                    $"Character '{value[i]}' (U+{code:X4}) at position {i} cannot be encoded as Code 128: " +
                    "only ASCII 32-126 (space through '~') is supported.",
                    nameof(value));
            values[i] = code - 32;
        }

        long checksum = StartB;
        for (int i = 0; i < values.Length; i++)
            checksum += (long)(i + 1) * values[i];
        checksum %= 103;

        var symbols = new List<int>(values.Length + 3) { StartB };
        symbols.AddRange(values);
        symbols.Add((int)checksum);
        symbols.Add(Stop);

        var widths = new List<int>(symbols.Count * 6);
        foreach (var symbol in symbols)
            foreach (var ch in Patterns[symbol])
                widths.Add(ch - '0');

        return [.. widths];
    }
}
