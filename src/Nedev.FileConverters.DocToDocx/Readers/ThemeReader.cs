using System.Text;
using Nedev.FileConverters.DocToDocx.Models;

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
                
                // Optional: Parse basic colors from the XML for internal use (e.g. preview)
                ParseThemeColors(document.Theme);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Failed to read theme: {ex.Message}");
        }
    }

    private static void ParseThemeColors(ThemeModel theme)
    {
        if (string.IsNullOrEmpty(theme.XmlContent)) return;

        // Very basic regex-based parsing to avoid heavy XML dependencies if not needed
        // Looking for <a:dk1>, <a:lt1>, <a:accent1>, etc.
        // Format: <a:clrScheme name="..."><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>...
        
        // This is a placeholder for more robust XML parsing if required.
        // For now, we'll just store the XML and let the writer embed it.
    }
}
