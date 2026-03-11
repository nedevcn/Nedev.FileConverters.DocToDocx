namespace Nedev.FileConverters.DocToDocx.Utils;

using Nedev.FileConverters.DocToDocx.Models;

/// <summary>
/// Helper for generating consistent style IDs for DOCX
/// </summary>
public static class StyleHelper
{
    /// <summary>
    /// Gets a consistent XML style ID for a paragraph style.
    /// Matches standard Word names to their standard IDs.
    /// </summary>
    public static string GetParagraphStyleId(int index, string? name)
    {
        if (index == 0 || string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase))
            return "Normal";

        if (index >= 1 && index <= 9)
            return $"Heading{index}";

        if (string.Equals(name, "Title", StringComparison.OrdinalIgnoreCase)) return "Title";
        if (string.Equals(name, "Subtitle", StringComparison.OrdinalIgnoreCase)) return "Subtitle";
        if (string.Equals(name, "Quote", StringComparison.OrdinalIgnoreCase)) return "Quote";
        if (string.Equals(name, "List Paragraph", StringComparison.OrdinalIgnoreCase)) return "ListParagraph";
        if (string.Equals(name, "No Spacing", StringComparison.OrdinalIgnoreCase)) return "NoSpacing";
        if (string.Equals(name, "Header", StringComparison.OrdinalIgnoreCase)) return "Header";
        if (string.Equals(name, "Footer", StringComparison.OrdinalIgnoreCase)) return "Footer";
        
        // Handle common variations
        if (!string.IsNullOrEmpty(name))
        {
            if (name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) && name.Length > 8)
            {
                if (int.TryParse(name.Substring(8), out int level) && level >= 1 && level <= 9)
                    return $"Heading{level}";
            }
            
            // Default to name without spaces
            return SanitizeStyleId(name.Replace(" ", string.Empty), $"Style{index}");
        }

        return $"Style{index}";
    }
    
    /// <summary>
    /// Gets a consistent XML style ID for a table style.
    /// </summary>
    public static string GetTableStyleId(int index, string? name)
    {
        if (index == 0 || string.Equals(name, "Normal Table", StringComparison.OrdinalIgnoreCase))
            return "TableNormal";

        if (index == 1 || string.Equals(name, "Table Grid", StringComparison.OrdinalIgnoreCase))
            return "TableGrid";

        if (!string.IsNullOrEmpty(name))
        {
            return SanitizeStyleId(name.Replace(" ", string.Empty), $"TableGrid{index}");
        }

        return $"TableGrid{index}";
    }

    /// <summary>
    /// Gets a consistent XML style ID for a character style.
    /// </summary>
    public static string GetCharacterStyleId(int index, string? name)
    {
        if (string.Equals(name, "Default Paragraph Font", StringComparison.OrdinalIgnoreCase))
            return "DefaultParagraphFont";

        if (!string.IsNullOrEmpty(name))
        {
            return SanitizeStyleId(name.Replace(" ", string.Empty), $"CharStyle{index}");
        }

        return $"CharStyle{index}";
    }

    /// <summary>
    /// Removes characters that are not valid in XML from a style display name
    /// and falls back when the resulting value would be empty.
    /// </summary>
    public static string GetSafeStyleName(string? name, string fallback)
    {
        var sanitized = SanitizeXmlString(name);
        return string.IsNullOrWhiteSpace(sanitized) ? fallback : sanitized.Trim();
    }

    private static string SanitizeStyleId(string? candidate, string fallback)
    {
        var sanitized = SanitizeXmlString(candidate);
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            return fallback;
        }

        sanitized = sanitized.Replace(" ", string.Empty);
        return string.IsNullOrWhiteSpace(sanitized) ? fallback : sanitized;
    }

    private static string SanitizeXmlString(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var sb = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '\uFFFD')
            {
                sb.Append(' ');
                continue;
            }

            if (c == '\t' || c == '\n' || c == '\r' ||
                (c >= 0x20 && c <= 0xD7FF) ||
                (c >= 0xE000 && c <= 0xFFFD))
            {
                sb.Append(c);
            }
            else if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                sb.Append(c);
                sb.Append(text[++i]);
            }
        }

        return sb.ToString();
    }
}
