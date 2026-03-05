using System.Security.Cryptography;
using System.Text;

namespace Nedev.DocToDocx.Writers;

public static class FontObfuscator
{
    public static byte[] ObfuscateFont(byte[] fontData, string fontKey)
    {
        if (fontData == null || fontData.Length < 32 || string.IsNullOrEmpty(fontKey))
        {
            return fontData;
        }

        // Convert the string font key to bytes.
        // The fontKey format is "{XXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}" usually.
        string guidString = fontKey.Trim('{', '}');
        if (!Guid.TryParse(guidString, out Guid guid))
        {
            return fontData;
        }

        // According to Open XML standard (ECMA-376 Part 4, 3.1.2)
        // 1. Get string value of GUID based font key ({...})
        // 2. Remove braces and hyphens. 
        // Note: the spec says "delete the { and } characters and the hyphens"
        // Then convert hex pairs to bytes, REVERSING the order of bytes?
        // Wait, standard procedure:
        // The GUID is represented as an array of 16 bytes.
        // In the byte array, the byte order is:
        // Reversed 1st 4 bytes, Reversed next 2 bytes, Reversed next 2 bytes, then remaining 8 bytes sequential.
        // This is exactly what Guid.ToByteArray() produces in .NET!
        
        byte[] guidBytes = guid.ToByteArray();
        
        // We need 32 bytes for the XOR key: the 16 byte GUID twice.
        // HOWEVER, the byte order for obfuscation is reversed.
        // Specifically, it's the reverse of the byte sequence of the GUID string where each pair of hex chars is a byte.
        // Example: if guid string is 61E4... then the bytes are 0x61, 0xE4, etc.
        // The spec says: The font key is the GUID string reversed, taking 2 characters at a time.
        
        byte[] obfuscationKey = new byte[32];
        string cleanGuid = guidString.Replace("-", "").ToUpperInvariant();
        
        // Reverse the pairs of characters
        byte[] keyBytes = new byte[16];
        for (int i = 0; i < 16; i++)
        {
            string hexPair = cleanGuid.Substring((15 - i) * 2, 2);
            keyBytes[i] = Convert.ToByte(hexPair, 16);
        }

        Buffer.BlockCopy(keyBytes, 0, obfuscationKey, 0, 16);
        Buffer.BlockCopy(keyBytes, 0, obfuscationKey, 16, 16);

        // Copy font data and XOR the first 32 bytes
        byte[] obfuscatedData = new byte[fontData.Length];
        Buffer.BlockCopy(fontData, 0, obfuscatedData, 0, fontData.Length);

        for (int i = 0; i < 32; i++)
        {
            obfuscatedData[i] ^= obfuscationKey[i];
        }

        return obfuscatedData;
    }
}
