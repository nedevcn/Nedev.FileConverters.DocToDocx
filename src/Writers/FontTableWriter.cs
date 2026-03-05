using System.Xml;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Writes the word/fontTable.xml file for DOCX
/// </summary>
public class FontTableWriter
{
    private readonly XmlWriter _writer;

    public FontTableWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    /// <summary>
    /// Writes the fontTable.xml content
    /// </summary>
    public void WriteFontTable(DocumentModel document, Dictionary<string, string> fontRelIds)
    {
        _writer.WriteStartDocument();
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        _writer.WriteStartElement("w", "fonts", wNs);
        _writer.WriteAttributeString("xmlns", "w", null, wNs);
        _writer.WriteAttributeString("xmlns", "r", null, rNs);

        foreach (var font in document.Styles.Fonts)
        {
            _writer.WriteStartElement("w", "font", wNs);
            _writer.WriteAttributeString("w", "name", wNs, font.Name);

            if (!string.IsNullOrEmpty(font.AltName))
            {
                _writer.WriteStartElement("w", "altName", wNs);
                _writer.WriteAttributeString("w", "val", wNs, font.AltName);
                _writer.WriteEndElement();
            }

            // Write pitch and family if needed
            if (font.Pitch > 0)
            {
                _writer.WriteStartElement("w", "pitch", wNs);
                _writer.WriteAttributeString("w", "val", wNs, GetPitchString(font.Pitch));
                _writer.WriteEndElement();
            }

            // Write embed relationship if font is embedded
            if (font.EmbeddedData != null && font.EmbeddedData.Length > 0 && fontRelIds.TryGetValue(font.Name, out var relId))
            {
                _writer.WriteStartElement("w", "embedRegular", wNs);
                _writer.WriteAttributeString("r", "id", rNs, relId);
                // The correct font key is derived from the font table's fontkey attribute,
                // but usually the relationship itself is enough for Word if unobfuscated,
                // or we provide a fontKey if obfuscated. We'll add fontKey later if using odttf.
                // We'll generate a GUID key for each font. We'll pass it in if needed.
                var fontKey = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
                // Actually, Word's fontKey is `{guid}`. 
                // Let's add the fontKey attribute
                _writer.WriteAttributeString("w", "fontKey", wNs, fontKey);
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); // w:font
        }

        _writer.WriteEndElement(); // w:fonts
        _writer.WriteEndDocument();
    }

    private string GetPitchString(int pitch)
    {
        // simplistic mapping based on MS-DOC pitch values
        return pitch switch
        {
            1 => "fixed",
            2 => "variable",
            _ => "default"
        };
    }
}
