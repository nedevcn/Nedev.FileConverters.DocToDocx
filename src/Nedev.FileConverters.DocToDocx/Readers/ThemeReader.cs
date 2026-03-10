using System.Text;
using System.Xml.Linq;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Extracts OOXML theme data from the "Theme" storage in legacy DOC files.
/// </summary>
public static class ThemeReader
{
    public static void Read(CfbReader cfb, DocumentModel document)
    {
        if (!cfb.HasStorage("Theme"))
            return;

        try
        {
            var themeStorage = cfb.GetStorage("Theme");
            if (themeStorage == null) return;

            // In some versions, the theme is in a stream named "Theme" within the "Theme" storage
            // In others, it might be "theme1.xml" if it's a direct copy
            var children = cfb.GetChildren(themeStorage);
            var themeEntry = children.FirstOrDefault(c => 
                c.Name.Equals("Theme", StringComparison.OrdinalIgnoreCase) ||
                c.Name.Equals("theme1.xml", StringComparison.OrdinalIgnoreCase));

            if (themeEntry != null)
            {
                var bytes = cfb.GetStreamBytes(themeEntry);
                document.Theme.XmlContent = Encoding.UTF8.GetString(bytes);

                ParseThemeMetadata(document.Theme);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read theme storage.", ex);
        }
    }

    public static void ParseThemeMetadata(ThemeModel theme)
    {
        if (string.IsNullOrEmpty(theme.XmlContent)) return;

        try
        {
            var document = XDocument.Parse(theme.XmlContent, LoadOptions.PreserveWhitespace);
            XNamespace aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

            theme.ColorMap.Clear();

            var clrScheme = document.Descendants(aNs + "clrScheme").FirstOrDefault();
            if (clrScheme != null)
            {
                foreach (var entry in clrScheme.Elements())
                {
                    var colorHex = ExtractThemeColorHex(entry, aNs);
                    if (!string.IsNullOrEmpty(colorHex))
                    {
                        theme.ColorMap[entry.Name.LocalName] = colorHex;
                    }
                }
            }

            var fontScheme = document.Descendants(aNs + "fontScheme").FirstOrDefault();
            if (fontScheme != null)
            {
                ApplyFontCollection(fontScheme.Element(aNs + "majorFont"), isMajor: true, theme, aNs);
                ApplyFontCollection(fontScheme.Element(aNs + "minorFont"), isMajor: false, theme, aNs);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to parse theme metadata; keeping raw theme XML only.", ex);
        }
    }

    private static void ApplyFontCollection(XElement? fontCollection, bool isMajor, ThemeModel theme, XNamespace aNs)
    {
        if (fontCollection == null)
            return;

        var latin = NormalizeFontName(fontCollection.Element(aNs + "latin")?.Attribute("typeface")?.Value);
        var eastAsia = NormalizeFontName(fontCollection.Element(aNs + "ea")?.Attribute("typeface")?.Value);
        var bidi = NormalizeFontName(fontCollection.Element(aNs + "cs")?.Attribute("typeface")?.Value);

        if (isMajor)
        {
            theme.MajorLatinFont = latin;
            theme.MajorEastAsiaFont = eastAsia;
            theme.MajorBidiFont = bidi;
        }
        else
        {
            theme.MinorLatinFont = latin;
            theme.MinorEastAsiaFont = eastAsia;
            theme.MinorBidiFont = bidi;
        }
    }

    private static string? ExtractThemeColorHex(XElement schemeEntry, XNamespace aNs)
    {
        var srgb = schemeEntry.Element(aNs + "srgbClr")?.Attribute("val")?.Value;
        if (IsHexColor(srgb))
            return srgb!.ToUpperInvariant();

        var sysClr = schemeEntry.Element(aNs + "sysClr");
        var lastClr = sysClr?.Attribute("lastClr")?.Value;
        if (IsHexColor(lastClr))
            return lastClr!.ToUpperInvariant();

        var val = sysClr?.Attribute("val")?.Value;
        if (IsHexColor(val))
            return val!.ToUpperInvariant();

        var scrgb = schemeEntry.Element(aNs + "scrgbClr");
        if (scrgb != null)
        {
            var r = ParseScrgbChannel(scrgb.Attribute("r")?.Value);
            var g = ParseScrgbChannel(scrgb.Attribute("g")?.Value);
            var b = ParseScrgbChannel(scrgb.Attribute("b")?.Value);
            if (r.HasValue && g.HasValue && b.HasValue)
                return $"{r.Value:X2}{g.Value:X2}{b.Value:X2}";
        }

        return null;
    }

    private static byte? ParseScrgbChannel(string? raw)
    {
        if (!int.TryParse(raw, out var value))
            return null;

        value = Math.Clamp(value, 0, 100000);
        return (byte)Math.Round(value * 255d / 100000d);
    }

    private static bool IsHexColor(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value!.Length != 6)
            return false;

        foreach (var ch in value)
        {
            if (!Uri.IsHexDigit(ch))
                return false;
        }

        return true;
    }

    private static string? NormalizeFontName(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }
}
