using System.Text;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Reads embedded fonts from the "FontTable" stream in legacy .doc files.
/// </summary>
public class FontReader
{
    private readonly byte[] _fontTableData;

    public FontReader(byte[] fontTableData)
    {
        _fontTableData = fontTableData;
    }

    public void ExtractFonts(StyleSheet styles)
    {
        if (_fontTableData == null || _fontTableData.Length < 4) return;

        using var ms = new MemoryStream(_fontTableData);
        using var reader = new BinaryReader(ms);

        try
        {
            // The FontTable stream is a series of font data blocks.
            // Each block typically starts with a length or a specific header.
            // Note: The structure of the "FontTable" stream can vary. 
            // Often it's a sequence of (length, fontName, data).
            
            while (ms.Position < ms.Length - 4)
            {
                int blockSize = reader.ReadInt32();
                if (blockSize <= 0 || ms.Position + blockSize > ms.Length) break;

                var blockData = reader.ReadBytes(blockSize);
                ProcessFontBlock(blockData, styles);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Error extracting embedded fonts: {ex.Message}");
        }
    }

    private void ProcessFontBlock(byte[] data, StyleSheet styles)
    {
        if (data.Length < 32) return;

        using var ms = new MemoryStream(data);
        using var reader = new BinaryReader(ms);

        // Word embedded font block header (simplified):
        // Offset 0: Font Name (varied encoding, often null-terminated)
        // Offset?: Font Data
        
        // This is a heuristic-based extraction because the "FontTable" stream
        // isn't as strictly documented as the main DOC structures.
        
        // Try to find the font name and the start of the 'sfnt' (TTF) data.
        // TTF files start with 0x00 0x01 0x00 0x00 or 'OTTO'.
        
        byte[] ttfSignature = { 0x00, 0x01, 0x00, 0x00 };
        byte[] ottoSignature = Encoding.ASCII.GetBytes("OTTO");
        
        int ttfOffset = FindSignature(data, ttfSignature);
        if (ttfOffset == -1) ttfOffset = FindSignature(data, ottoSignature);
        
        if (ttfOffset != -1)
        {
            // Extract font name from the data before the TTF signature
            string fontName = ExtractFontName(data, ttfOffset);
            if (!string.IsNullOrEmpty(fontName))
            {
                var font = styles.Fonts.FirstOrDefault(f => f.Name.Equals(fontName, StringComparison.OrdinalIgnoreCase));
                if (font != null)
                {
                    byte[] fontData = new byte[data.Length - ttfOffset];
                    Buffer.BlockCopy(data, ttfOffset, fontData, 0, fontData.Length);
                    font.EmbeddedData = fontData;
                }
            }
        }
    }

    private int FindSignature(byte[] data, byte[] signature)
    {
        for (int i = 0; i <= data.Length - signature.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < signature.Length; j++)
            {
                if (data[i + j] != signature[j])
                {
                    match = false;
                    break;
                }
            }
            if (match) return i;
        }
        return -1;
    }

    private string ExtractFontName(byte[] data, int endOffset)
    {
        // Search backwards for the font name
        // It's usually a null-terminated Unicode or ANSI string.
        
        // Look for the last null terminator before the TTF data
        int i = endOffset - 1;
        while (i > 0 && data[i] == 0) i--;
        
        int endOfName = i + 1;
        int startOfName = i;
        
        while (startOfName > 0 && data[startOfName] != 0) startOfName--;
        if (data[startOfName] == 0) startOfName++;
        
        if (endOfName > startOfName)
        {
            try
            {
                // Try Unicode first (standard for modern Word)
                string name = Encoding.Unicode.GetString(data, startOfName, endOfName - startOfName);
                if (IsLegalFontName(name)) return name;
                
                // Fallback to ANSI
                name = Encoding.ASCII.GetString(data, startOfName, endOfName - startOfName);
                if (IsLegalFontName(name)) return name;
            }
            catch { }
        }
        
        return string.Empty;
    }

    private bool IsLegalFontName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return false;
        foreach (char c in name)
        {
            if (char.IsControl(c)) return false;
        }
        return name.Length > 2 && name.Length < 64;
    }
}
