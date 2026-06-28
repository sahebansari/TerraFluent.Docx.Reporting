namespace TerraFluent.Docx.Reporting.Ooxml;

internal static class ImageReader
{
    // Returns (widthPx, heightPx). Falls back to (0,0) if unreadable.
    public static (int width, int height) GetPixelDimensions(string filePath)
    {
        try
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            using var fs = File.OpenRead(filePath);
            return ext switch
            {
                ".png"          => ReadPng(fs),
                ".jpg" or ".jpeg" => ReadJpeg(fs),
                _               => (0, 0)
            };
        }
        catch
        {
            return (0, 0);
        }
    }

    public static (int width, int height) GetPixelDimensions(byte[] bytes, string fileName)
    {
        try
        {
            using var ms = new MemoryStream(bytes);
            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            return ext switch
            {
                ".png" => ReadPng(ms),
                ".jpg" or ".jpeg" => ReadJpeg(ms),
                _ => (0, 0)
            };
        }
        catch
        {
            return (0, 0);
        }
    }

    public static string ContentType(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".png"            => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif"            => "image/gif",
            _                 => "application/octet-stream"
        };
    }

    // PNG: signature = 8 bytes, then IHDR chunk: 4-byte length, 4-byte "IHDR",
    //      4-byte width (big-endian), 4-byte height (big-endian)
    private static (int, int) ReadPng(Stream stream)
    {
        var buf = new byte[24];
        if (stream.Read(buf, 0, buf.Length) < 24) return (0, 0);
        // Check PNG signature
        if (buf[0] != 137 || buf[1] != 80 || buf[2] != 78 || buf[3] != 71) return (0, 0);
        int w = (buf[16] << 24) | (buf[17] << 16) | (buf[18] << 8) | buf[19];
        int h = (buf[20] << 24) | (buf[21] << 16) | (buf[22] << 8) | buf[23];
        return (w, h);
    }

    // JPEG: SOI=FFD8, then scan for SOF0/SOF1/SOF2 (FFC0/FFC1/FFC2) markers
    private static (int, int) ReadJpeg(Stream stream)
    {
        var two = new byte[2];

        if (stream.Read(two, 0, 2) < 2 || two[0] != 0xFF || two[1] != 0xD8) return (0, 0);

        var lenBuf = new byte[2];
        var sofBuf = new byte[4];

        while (stream.Position < stream.Length - 1)
        {
            // Find 0xFF marker
            int b = stream.ReadByte();
            if (b != 0xFF) continue;
            int marker = stream.ReadByte();

            if (marker is >= 0xC0 and <= 0xC3)
            {
                // SOF marker: 2-byte length, 1-byte precision, 2-byte height, 2-byte width
                stream.Seek(3, SeekOrigin.Current);
                if (stream.Read(sofBuf, 0, 4) < 4) break;
                int h = (sofBuf[0] << 8) | sofBuf[1];
                int w = (sofBuf[2] << 8) | sofBuf[3];
                return (w, h);
            }

            // Skip this segment
            if (stream.Read(lenBuf, 0, 2) < 2) break;
            int segLen = (lenBuf[0] << 8) | lenBuf[1];
            if (segLen < 2) break;
            stream.Seek(segLen - 2, SeekOrigin.Current);
        }
        return (0, 0);
    }
}
