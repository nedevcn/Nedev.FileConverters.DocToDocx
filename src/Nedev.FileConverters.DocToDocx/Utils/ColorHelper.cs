using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Utils;

/// <summary>
/// Color conversion utilities for DOC → DOCX
/// </summary>
public static class ColorHelper
{
    /// <summary>
    /// Word ICO color index (0-16) to RGB mapping
    /// Reference: MS-DOC 2.9.98 ico
    /// </summary>
    private static readonly uint[] IcoToRgbTable = new uint[]
    {
        0x000000, // 0 = Auto (black)
        0x000000, // 1 = Black
        0x0000FF, // 2 = Blue
        0x00FFFF, // 3 = Cyan
        0x00FF00, // 4 = Green
        0xFF00FF, // 5 = Magenta
        0xFF0000, // 6 = Red
        0xFFFF00, // 7 = Yellow
        0xFFFFFF, // 8 = White
        0x00008B, // 9 = DarkBlue
        0x008B8B, // 10 = DarkCyan
        0x006400, // 11 = DarkGreen
        0x8B008B, // 12 = DarkMagenta
        0x8B0000, // 13 = DarkRed
        0x808000, // 14 = DarkYellow (Olive)
        0xA9A9A9, // 15 = DarkGray
        0xD3D3D3, // 16 = LightGray
    };

    /// <summary>
    /// Word highlight color index to OOXML highlight name mapping
    /// </summary>
    private static readonly string[] HighlightNames = new string[]
    {
        "none",        // 0
        "black",       // 1
        "blue",        // 2
        "cyan",        // 3
        "green",       // 4
        "magenta",     // 5
        "red",         // 6
        "yellow",      // 7
        "white",       // 8
        "darkBlue",    // 9
        "darkCyan",    // 10
        "darkGreen",   // 11
        "darkMagenta", // 12
        "darkRed",     // 13
        "darkYellow",  // 14
        "darkGray",    // 15
        "lightGray",   // 16
    };

    /// <summary>
    /// OOXML theme color names mapping from MS-DOC theme index
    /// </summary>
    private static readonly string[] ThemeColorSchemeNames = new string[]
    {
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", 
        "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"
    };

    /// <summary>
    /// WordprocessingML theme color names for w:color/@w:themeColor.
    /// </summary>
    private static readonly string[] ThemeColorWordNames = new string[]
    {
        "dark1", "light1", "dark2", "light2", "accent1", "accent2",
        "accent3", "accent4", "accent5", "accent6", "hyperlink", "followedHyperlink"
    };

    /// <summary>
    /// Converts a Word ICO color index to a 6-digit hex RGB string (e.g. "FF0000")
    /// </summary>
    public static string IcoToRgbHex(int ico)
    {
        if (ico < 0 || ico >= IcoToRgbTable.Length)
            return "000000";
        return IcoToRgbTable[ico].ToString("X6");
    }

    /// <summary>
    /// Converts a Word ICO color index to an RGB uint value
    /// </summary>
    public static uint IcoToRgb(int ico)
    {
        if (ico < 0 || ico >= IcoToRgbTable.Length)
            return 0;
        return IcoToRgbTable[ico];
    }

    /// <summary>
    /// Gets the OOXML highlight color name for a highlight index
    /// </summary>
    public static string GetHighlightName(int index)
    {
        if (index < 0 || index >= HighlightNames.Length)
            return "none";
        return HighlightNames[index];
    }

    /// <summary>
    /// Converts a 24-bit RGB integer to a 6-digit hex string
    /// </summary>
    public static string RgbToHex(uint rgb)
    {
        // MS-DOC stores colors as COLORREF (0x00BBGGRR)
        // OOXML expects RRGGBB
        var r = rgb & 0xFF;
        var g = (rgb >> 8) & 0xFF;
        var b = (rgb >> 16) & 0xFF;
        return $"{r:X2}{g:X2}{b:X2}";
    }

    /// <summary>
    /// Determines if a color value is likely an ICO index (0-16) or a direct RGB value
    /// Values 0-16 are ambiguous, but values > 16 are definitely RGB
    /// </summary>
    public static bool IsLikelyIcoIndex(int color)
    {
        return color >= 0 && color <= 16;
    }

    /// <summary>
    /// Converts a color value to a hex string, auto-detecting ICO vs RGB.
    /// For values 0-16, treats as ICO index. For values > 16, treats as direct RGB.
    /// </summary>
    public static string ColorToHex(int color)
    {
        if (color == 0) return "auto";
        if (IsLikelyIcoIndex(color))
            return IcoToRgbHex(color);
            
        // Check for theme color (bit 24 set to 1)
        if ((color & 0x01000000) != 0)
        {
            // For now, we return "auto" or handle it via GetThemeColorName if we want the name
            return "auto"; 
        }

        // Convert COLORREF to RRGGBB
        var r = color & 0xFF;
        var g = (color >> 8) & 0xFF;
        var b = (color >> 16) & 0xFF;
        return $"{r:X2}{g:X2}{b:X2}";
    }

    /// <summary>
    /// Gets the theme color name if the color is a theme color.
    /// Returns null if not a theme color.
    /// </summary>
    public static string? GetThemeColorName(int color)
    {
        if ((color & 0x01000000) != 0)
        {
            int index = color & 0xFF;
            if (index >= 0 && index < ThemeColorWordNames.Length)
                return ThemeColorWordNames[index];
        }
        return null;
    }

    /// <summary>
    /// Gets the DrawingML theme scheme color name used by theme1.xml.
    /// </summary>
    public static string? GetThemeSchemeColorName(int color)
    {
        if ((color & 0x01000000) != 0)
        {
            int index = color & 0xFF;
            if (index >= 0 && index < ThemeColorSchemeNames.Length)
                return ThemeColorSchemeNames[index];
        }

        return null;
    }

    /// <summary>
    /// Resolves a theme-backed color to a concrete RGB hex value using the parsed theme metadata.
    /// </summary>
    public static string? ResolveThemeColorHex(int color, ThemeModel? theme)
    {
        if (theme == null || theme.ColorMap.Count == 0)
            return null;

        var schemeName = GetThemeSchemeColorName(color);
        if (schemeName == null)
            return null;

        return theme.ColorMap.TryGetValue(schemeName, out var hex) ? hex : null;
    }
}
