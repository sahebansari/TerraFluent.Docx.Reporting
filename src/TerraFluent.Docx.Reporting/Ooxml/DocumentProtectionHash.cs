using System.Security.Cryptography;
using System.Text;

namespace TerraFluent.Docx.Reporting.Ooxml;

/// <summary>
/// Computes the <c>w:documentProtection</c> password verifier per ISO/IEC 29500-1 §17.15.1.28:
/// the password is first reduced to Word's legacy 32-bit key, rendered as an 8-character hex
/// string, then hashed as SHA-512(salt + UTF-16LE(key)) and iterated
/// SHA-512(previous + iteration-index as 4-byte little-endian) for the spin count.
/// </summary>
internal static class DocumentProtectionHash
{
    public const int SpinCount = 100000;

    private static readonly int[] InitialCodeArray =
    [
        0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E,
        0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3
    ];

    private static readonly int[,] EncryptionMatrix =
    {
        { 0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09 },
        { 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF },
        { 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0 },
        { 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40 },
        { 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5 },
        { 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A },
        { 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9 },
        { 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0 },
        { 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC },
        { 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10 },
        { 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168 },
        { 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C },
        { 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD },
        { 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC },
        { 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4 }
    };

    public static byte[] GenerateSalt()
    {
        var salt = new byte[16];
        using var rng = RandomNumberGenerator.Create();
        rng.GetBytes(salt);
        return salt;
    }

    public static byte[] ComputeHash(string password, byte[] salt, int spinCount = SpinCount)
    {
        var keyBytes = Encoding.Unicode.GetBytes(LegacyKeyHex(password));

        using var sha512 = SHA512.Create();

        var buffer = new byte[salt.Length + keyBytes.Length];
        salt.CopyTo(buffer, 0);
        keyBytes.CopyTo(buffer, salt.Length);
        var hash = sha512.ComputeHash(buffer);

        // Each round appends the zero-based iteration index as a 4-byte little-endian integer.
        var iterated = new byte[hash.Length + 4];
        for (int i = 0; i < spinCount; i++)
        {
            hash.CopyTo(iterated, 0);
            iterated[hash.Length]     = (byte)i;
            iterated[hash.Length + 1] = (byte)(i >> 8);
            iterated[hash.Length + 2] = (byte)(i >> 16);
            iterated[hash.Length + 3] = (byte)(i >> 24);
            hash = sha512.ComputeHash(iterated);
        }

        return hash;
    }

    // Word's legacy binary-format password key: 15-char max, per-character low byte (high byte
    // when the low byte is zero), a matrix-XOR high word and a rotate-XOR low word, combined and
    // written as the byte-reversed 8-digit hex string.
    internal static string LegacyKeyHex(string password)
    {
        if (password.Length > 15)
            password = password.Substring(0, 15);

        var chars = new byte[password.Length];
        for (int i = 0; i < password.Length; i++)
        {
            int code = password[i];
            chars[i] = (byte)(code & 0x00FF);
            if (chars[i] == 0)
                chars[i] = (byte)((code & 0xFF00) >> 8);
        }

        // The matrix rows correspond to the last 15 character positions, so a shorter password
        // starts partway down: row = 15 - length + i.
        int high = InitialCodeArray[chars.Length - 1];
        for (int i = 0; i < chars.Length; i++)
        {
            int row = 15 - chars.Length + i;
            for (int bit = 0; bit < 7; bit++)
            {
                if ((chars[i] & (1 << bit)) != 0)
                    high ^= EncryptionMatrix[row, bit];
            }
        }

        int low = 0;
        for (int i = chars.Length - 1; i >= 0; i--)
            low = (((low >> 14) & 0x0001) | ((low << 1) & 0x7FFF)) ^ chars[i];
        low = (((low >> 14) & 0x0001) | ((low << 1) & 0x7FFF)) ^ chars.Length ^ 0xCE4B;

        // Byte-reverse the combined key and render as 8 zero-padded uppercase hex digits. (The
        // widely-copied MSDN sample skips the zero padding, which breaks any password whose key
        // contains a byte below 0x10 - e.g. "24209903" - verified against Word itself.)
        int combined = (high << 16) + low;
        var sb = new StringBuilder(8);
        for (int i = 0; i < 4; i++)
            sb.Append(((combined >> (i * 8)) & 0xFF).ToString("X2"));
        return sb.ToString();
    }
}
