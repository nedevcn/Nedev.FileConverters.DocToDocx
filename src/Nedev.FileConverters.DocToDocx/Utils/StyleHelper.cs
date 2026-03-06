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
            return name.Replace(" ", "");
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
            return name.Replace(" ", "");
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
            return name.Replace(" ", "");
        }
        
        return $"CharStyle{index}";
    }
}
