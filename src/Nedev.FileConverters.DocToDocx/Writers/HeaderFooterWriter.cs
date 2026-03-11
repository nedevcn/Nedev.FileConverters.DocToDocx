using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// Writes header and footer XML files for DOCX
/// </summary>
public class HeaderFooterWriter
{
    private readonly XmlWriter _writer;
    private const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public HeaderFooterWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    /// <summary>
    /// Writes a header XML file
    /// </summary>
    public void WriteHeader(
        HeaderFooterModel header,
        DocumentModel document,
        IReadOnlyDictionary<int, string>? imageRelationshipIds = null,
        IReadOnlyDictionary<string, string>? oleRelationshipIds = null)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "hdr", Wns);
        _writer.WriteAttributeString("xmlns", "w", null, Wns);
        WriteRootNamespaces();

        // Write header content
        WriteHeaderFooterContent(header, document, isHeader: true, imageRelationshipIds, oleRelationshipIds);

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }

    /// <summary>
    /// Writes a footer XML file
    /// </summary>
    public void WriteFooter(
        HeaderFooterModel footer,
        DocumentModel document,
        IReadOnlyDictionary<int, string>? imageRelationshipIds = null,
        IReadOnlyDictionary<string, string>? oleRelationshipIds = null)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "ftr", Wns);
        _writer.WriteAttributeString("xmlns", "w", null, Wns);
        WriteRootNamespaces();

        // Write footer content
        WriteHeaderFooterContent(footer, document, isHeader: false, imageRelationshipIds, oleRelationshipIds);

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }

    /// <summary>
    /// Writes header/footer content (paragraphs)
    /// </summary>
    private void WriteHeaderFooterContent(
        HeaderFooterModel headerFooter,
        DocumentModel document,
        bool isHeader,
        IReadOnlyDictionary<int, string>? imageRelationshipIds,
        IReadOnlyDictionary<string, string>? oleRelationshipIds)
    {
        if (!HeaderFooterContentHelper.HasUsableContent(headerFooter))
        {
            WriteDefaultHeaderFooterParagraph();
            return;
        }

        if (headerFooter.Paragraphs != null && headerFooter.Paragraphs.Count > 0)
        {
            var docWriter = new DocumentWriter(
                _writer,
                new DocumentWriterOptions
                {
                    EnableHyperlinks = false
                })
                .BindDocumentContext(document, null, imageRelationshipIds, oleRelationshipIds);

            foreach (var paragraph in headerFooter.Paragraphs)
            {
                docWriter.WriteParagraph(paragraph);
            }
        }
        else if (!string.IsNullOrEmpty(headerFooter.Text))
        {
            // Simple text content fallback
            WriteSimpleParagraph(headerFooter.Text, isHeader);
        }
        else
        {
            // Empty content placeholder
            WriteDefaultHeaderFooterParagraph();
        }
    }

    /// <summary>
    /// Writes a simple paragraph with text
    /// </summary>
    private void WriteSimpleParagraph(string text, bool isHeader)
    {
        _writer.WriteStartElement("w", "p", Wns);

        // Paragraph properties - center aligned for headers/footers by default
        _writer.WriteStartElement("w", "pPr", Wns);
        _writer.WriteStartElement("w", "jc", Wns);
        _writer.WriteAttributeString("w", "val", null, "center");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Run with text
        _writer.WriteStartElement("w", "r", Wns);

        // Run properties
        _writer.WriteStartElement("w", "rPr", Wns);
        _writer.WriteStartElement("w", "rStyle", Wns);
        _writer.WriteAttributeString("w", "val", null, isHeader ? "Header" : "Footer");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Text
        _writer.WriteStartElement("w", "t", Wns);
        WriteTextWithEscape(text);
        _writer.WriteEndElement();

        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p
    }

    /// <summary>
    /// Writes a default paragraph for empty headers/footers
    /// </summary>
    private void WriteDefaultHeaderFooterParagraph()
    {
        _writer.WriteStartElement("w", "p", Wns);

        // Paragraph properties
        _writer.WriteStartElement("w", "pPr", Wns);
        _writer.WriteStartElement("w", "jc", Wns);
        _writer.WriteAttributeString("w", "val", null, "center");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Empty paragraph - no run element needed
        _writer.WriteEndElement();
    }

    /// <summary>
    /// Writes a paragraph with page number field
    /// </summary>
    public void WritePageNumberParagraph()
    {
        _writer.WriteStartElement("w", "p", Wns);

        // Paragraph properties
        _writer.WriteStartElement("w", "pPr", Wns);
        _writer.WriteStartElement("w", "jc", Wns);
        _writer.WriteAttributeString("w", "val", null, "center");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Page number field
        _writer.WriteStartElement("w", "r", Wns);

        // Field begin
        _writer.WriteStartElement("w", "fldChar", Wns);
        _writer.WriteAttributeString("w", "fldCharType", null, "begin");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Field code
        _writer.WriteStartElement("w", "r", Wns);
        _writer.WriteStartElement("w", "instrText", Wns);
        _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        _writer.WriteString("PAGE");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Field separator
        _writer.WriteStartElement("w", "r", Wns);
        _writer.WriteStartElement("w", "fldChar", Wns);
        _writer.WriteAttributeString("w", "fldCharType", null, "separate");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Field result (page number)
        _writer.WriteStartElement("w", "r", Wns);
        _writer.WriteStartElement("w", "t", Wns);
        _writer.WriteString("1");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Field end
        _writer.WriteStartElement("w", "r", Wns);
        _writer.WriteStartElement("w", "fldChar", Wns);
        _writer.WriteAttributeString("w", "fldCharType", null, "end");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        _writer.WriteEndElement(); // w:p
    }

    /// <summary>
    /// Writes text with proper XML escaping
    /// </summary>
    private void WriteTextWithEscape(string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        var cleanedText = DocumentWriter.SanitizeXmlString(text);
        
        // Check if text needs space preservation
        if (cleanedText.StartsWith(' ') || cleanedText.EndsWith(' ') || cleanedText.Contains("  "))
        {
            _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        }

        _writer.WriteString(cleanedText);
    }

    private void WriteRootNamespaces()
    {
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        _writer.WriteAttributeString("xmlns", "wp", null, "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("xmlns", "pic", null, "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteAttributeString("xmlns", "wps", null, "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        _writer.WriteAttributeString("xmlns", "v", null, "urn:schemas-microsoft-com:vml");
        _writer.WriteAttributeString("xmlns", "o", null, "urn:schemas-microsoft-com:office:office");
        _writer.WriteAttributeString("xmlns", "c", null, "http://schemas.openxmlformats.org/drawingml/2006/chart");
    }
}

/// <summary>
/// Manages header/footer relationships and references
/// </summary>
public class HeaderFooterRelationshipManager
{
    private readonly Dictionary<HeaderFooterType, string> _relationshipIds = new();
    private int _nextId = 1;

    /// <summary>
    /// Gets or creates a relationship ID for a header/footer
    /// </summary>
    public string GetRelationshipId(HeaderFooterType type)
    {
        if (!_relationshipIds.TryGetValue(type, out var id))
        {
            id = $"rId{_nextId++}";
            _relationshipIds[type] = id;
        }
        return id;
    }

    /// <summary>
    /// Gets the file name for a header/footer type
    /// </summary>
    public string GetFileName(HeaderFooterType type)
    {
        return type switch
        {
            HeaderFooterType.HeaderFirst => "header1.xml",
            HeaderFooterType.HeaderOdd => "header2.xml",
            HeaderFooterType.HeaderEven => "header3.xml",
            HeaderFooterType.FooterFirst => "footer1.xml",
            HeaderFooterType.FooterOdd => "footer2.xml",
            HeaderFooterType.FooterEven => "footer3.xml",
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };
    }

    /// <summary>
    /// Gets the relationship type for a header/footer
    /// </summary>
    public string GetRelationshipType(HeaderFooterType type)
    {
        return type switch
        {
            HeaderFooterType.HeaderFirst or HeaderFooterType.HeaderOdd or HeaderFooterType.HeaderEven
                => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
            HeaderFooterType.FooterFirst or HeaderFooterType.FooterOdd or HeaderFooterType.FooterEven
                => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };
    }

    /// <summary>
    /// Gets all registered header/footer types
    /// </summary>
    public IEnumerable<HeaderFooterType> GetRegisteredTypes()
    {
        return _relationshipIds.Keys;
    }

    /// <summary>
    /// Checks if a header/footer type is registered
    /// </summary>
    public bool IsRegistered(HeaderFooterType type)
    {
        return _relationshipIds.ContainsKey(type);
    }
}
