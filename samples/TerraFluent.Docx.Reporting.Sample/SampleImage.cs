using System.IO.Compression;

internal static class SampleImage
{
    public static void CreateGradientPng(string path, int width, int height)
    {
        using var file = new MemoryStream();

        file.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        WriteChunk(file, "IHDR", [
            (byte)(width >> 24), (byte)(width >> 16), (byte)(width >> 8), (byte)width,
            (byte)(height >> 24), (byte)(height >> 16), (byte)(height >> 8), (byte)height,
            8, 2, 0, 0, 0
        ]);

        var raw = new byte[height * (1 + width * 3)];
        for (int y = 0; y < height; y++)
        {
            int row = y * (1 + width * 3);
            raw[row] = 0;
            for (int x = 0; x < width; x++)
            {
                raw[row + 1 + x * 3] = (byte)(x * 255 / (width - 1));
                raw[row + 1 + x * 3 + 1] = (byte)(y * 255 / (height - 1));
                raw[row + 1 + x * 3 + 2] = 128;
            }
        }

        WriteChunk(file, "IDAT", ZlibCompress(raw));
        WriteChunk(file, "IEND", []);
        File.WriteAllBytes(path, file.ToArray());
    }

    private static void WriteChunk(MemoryStream ms, string type, byte[] data)
    {
        var len = BitConverter.GetBytes(data.Length);
        if (BitConverter.IsLittleEndian) Array.Reverse(len);
        ms.Write(len);

        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        ms.Write(typeBytes);
        ms.Write(data);

        var crcBuffer = new byte[4 + data.Length];
        typeBytes.CopyTo(crcBuffer, 0);
        data.CopyTo(crcBuffer, 4);

        var crcValue = BitConverter.GetBytes(Crc32(crcBuffer));
        if (BitConverter.IsLittleEndian) Array.Reverse(crcValue);
        ms.Write(crcValue);
    }

    private static byte[] ZlibCompress(byte[] data)
    {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);

        using (var deflate = new DeflateStream(output, CompressionMode.Compress, true))
            deflate.Write(data, 0, data.Length);

        var adler = BitConverter.GetBytes(Adler32(data));
        if (BitConverter.IsLittleEndian) Array.Reverse(adler);
        output.Write(adler);
        return output.ToArray();
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }
        return ~crc;
    }

    private static uint Adler32(byte[] data)
    {
        uint s1 = 1;
        uint s2 = 0;
        foreach (var b in data)
        {
            s1 = (s1 + b) % 65521;
            s2 = (s2 + s1) % 65521;
        }
        return (s2 << 16) | s1;
    }
}
