using System.Xml;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Writes relationships files for DOCX
/// </summary>
public class RelationshipsWriter
{
    private readonly XmlWriter _writer;
    
    public RelationshipsWriter(XmlWriter writer)
    {
        _writer = writer;
    }
    
    private static string GetImageExtension(ImageType type)
    {
        return type switch
        {
            ImageType.Png => ".png",
            ImageType.Jpeg => ".jpg",
            ImageType.Gif => ".gif",
            ImageType.Emf => ".emf",
            ImageType.Wmf => ".wmf",
            ImageType.Dib => ".bmp",
            ImageType.Tiff => ".tiff",
            _ => ".png"
        };
    }
    
    /// <summary>
    /// Writes the main .rels file
    /// </summary>
    public void WriteMainRelationships()
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
        
        WriteRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "word/document.xml");
        WriteRelationship("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
        WriteRelationship("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
        
        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }
    
    /// <summary>
    /// Writes document relationships
    /// </summary>
    public void WriteDocumentRelationships(DocumentModel document)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
        
        var ids = ComputeRelationshipIds(document);
        
        WriteRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
        
        // Settings (always present)
        WriteRelationship($"rId{ids.SettingsRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml");
        
        // Font table relationship
        if (ids.FontTableRId > 0)
        {
            WriteRelationship($"rId{ids.FontTableRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable", "fontTable.xml");
        }
        
        // Theme relationship
        if (ids.ThemeRId > 0)
        {
            WriteRelationship($"rId{ids.ThemeRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml");
        }
        
        // Image relationships
        for (int i = 0; i < document.Images.Count; i++)
        {
            var extension = GetImageExtension(document.Images[i].Type);
            WriteRelationship($"rId{ids.FirstImageRId + i}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", $"media/image{i + 1}{extension}");
        }

        // Chart relationships (one per chart part, if any)
        if (document.Charts.Count > 0 && ids.FirstChartRId > 0)
        {
            for (int i = 0; i < document.Charts.Count; i++)
            {
                WriteRelationship(
                    $"rId{ids.FirstChartRId + i}",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                    $"charts/chart{i + 1}.xml");
            }
        }
        
        // OLE relationships
        if (document.OleObjects.Count > 0 && ids.FirstOleRId > 0)
        {
            for (int i = 0; i < document.OleObjects.Count; i++)
            {
                WriteRelationship(
                    $"rId{ids.FirstOleRId + i}",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                    $"embeddings/oleObject{i + 1}.bin");
            }
        }
        
        // Numbering relationship
        if (ids.NumberingRId > 0)
        {
            WriteRelationship($"rId{ids.NumberingRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml");
        }
        
        // Footnotes relationship
        if (ids.FootnotesRId > 0)
        {
            WriteRelationship($"rId{ids.FootnotesRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", "footnotes.xml");
        }
        
        // Endnotes relationship
        if (ids.EndnotesRId > 0)
        {
            WriteRelationship($"rId{ids.EndnotesRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", "endnotes.xml");
        }
        
        // VBA Project relationship
        if (ids.VbaProjectRId > 0)
        {
            WriteRelationship($"rId{ids.VbaProjectRId}", "http://schemas.microsoft.com/office/2006/relationships/vbaProject", "vbaProject.bin");
        }
        
        // Header relationships (up to three: first/odd/even)
        if (ids.HeaderFirstRId > 0)
        {
            WriteRelationship($"rId{ids.HeaderFirstRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", "header1.xml");
        }
        if (ids.HeaderOddRId > 0)
        {
            WriteRelationship($"rId{ids.HeaderOddRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", "header2.xml");
        }
        if (ids.HeaderEvenRId > 0)
        {
            WriteRelationship($"rId{ids.HeaderEvenRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", "header3.xml");
        }
        
        // Footer relationships (up to three: first/odd/even)
        if (ids.FooterFirstRId > 0)
        {
            WriteRelationship($"rId{ids.FooterFirstRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footer1.xml");
        }
        if (ids.FooterOddRId > 0)
        {
            WriteRelationship($"rId{ids.FooterOddRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footer2.xml");
        }
        if (ids.FooterEvenRId > 0)
        {
            WriteRelationship($"rId{ids.FooterEvenRId}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footer3.xml");
        }

        // Hyperlink relationships (external)
        WriteHyperlinkRelationships(document);
        
        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }
    
    /// <summary>
    /// Computes all relationship IDs for a document. Shared by DocumentWriter to ensure consistency.
    /// </summary>
    public static DocumentRelationshipIds ComputeRelationshipIds(DocumentModel document)
    {
        var ids = new DocumentRelationshipIds();
        var nextId = 2; // rId1 = styles
        
        ids.SettingsRId = nextId++;
        
        if (document.Styles.Fonts.Any(f => f.EmbeddedData != null))
            ids.FontTableRId = nextId++;
        
        if (!string.IsNullOrEmpty(document.Theme.XmlContent))
            ids.ThemeRId = nextId++;
        
        ids.FirstImageRId = nextId;
        nextId += document.Images.Count;

        // Reserve a contiguous block of relationship IDs for charts after images.
        // Charts are emitted as separate parts under word/charts/chartN.xml.
        if (document.Charts.Count > 0)
        {
            ids.FirstChartRId = nextId;
            nextId += document.Charts.Count;
        }
        
        if (document.OleObjects.Count > 0)
        {
            ids.FirstOleRId = nextId;
            nextId += document.OleObjects.Count;
        }
        
        bool hasNumbering = document.Paragraphs.Any(p => p.ListFormatId > 0) || document.NumberingDefinitions.Count > 0;
        if (hasNumbering)
            ids.NumberingRId = nextId++;
        
        if (document.Footnotes.Count > 0)
            ids.FootnotesRId = nextId++;
        
        if (document.Endnotes.Count > 0)
            ids.EndnotesRId = nextId++;
        
        if (document.VbaProject != null)
            ids.VbaProjectRId = nextId++;
        
        // Header/footer parts: allocate distinct IDs per type if present
        bool hasHeaderFirst = document.HeadersFooters.Headers.Any(h => h.Type == HeaderFooterType.HeaderFirst);
        bool hasHeaderOdd = document.HeadersFooters.Headers.Any(h => h.Type == HeaderFooterType.HeaderOdd);
        bool hasHeaderEven = document.HeadersFooters.Headers.Any(h => h.Type == HeaderFooterType.HeaderEven);

        if (hasHeaderFirst)
            ids.HeaderFirstRId = nextId++;
        if (hasHeaderOdd)
            ids.HeaderOddRId = nextId++;
        if (hasHeaderEven)
            ids.HeaderEvenRId = nextId++;

        bool hasFooterFirst = document.HeadersFooters.Footers.Any(f => f.Type == HeaderFooterType.FooterFirst);
        bool hasFooterOdd = document.HeadersFooters.Footers.Any(f => f.Type == HeaderFooterType.FooterOdd);
        bool hasFooterEven = document.HeadersFooters.Footers.Any(f => f.Type == HeaderFooterType.FooterEven);

        if (hasFooterFirst)
            ids.FooterFirstRId = nextId++;
        if (hasFooterOdd)
            ids.FooterOddRId = nextId++;
        if (hasFooterEven)
            ids.FooterEvenRId = nextId++;

        // Backward-compatible aggregate IDs (not used for relationships anymore but
        // kept in case other code relies on them).
        ids.HeaderRId = ids.HeaderOddRId != 0
            ? ids.HeaderOddRId
            : (ids.HeaderFirstRId != 0 ? ids.HeaderFirstRId : ids.HeaderEvenRId);

        ids.FooterRId = ids.FooterOddRId != 0
            ? ids.FooterOddRId
            : (ids.FooterFirstRId != 0 ? ids.FooterFirstRId : ids.FooterEvenRId);

        ids.LastUsedRId = nextId - 1;
        return ids;
    }
    
    private void WriteRelationship(string id, string type, string target)
    {
        _writer.WriteStartElement("Relationship");
        _writer.WriteAttributeString("Id", id);
        _writer.WriteAttributeString("Type", type);
        _writer.WriteAttributeString("Target", target);
        _writer.WriteEndElement();
    }

    private void WriteHyperlinkRelationships(DocumentModel document)
    {
        // Collect unique external hyperlinks that have relationship IDs assigned
        var hyperlinks = document.Hyperlinks
            .Where(h => h.IsExternal && !string.IsNullOrEmpty(h.Url) && !string.IsNullOrEmpty(h.RelationshipId))
            .GroupBy(h => h.RelationshipId)
            .Select(g => g.First());

        foreach (var link in hyperlinks)
        {
            _writer.WriteStartElement("Relationship");
            _writer.WriteAttributeString("Id", link.RelationshipId!);
            _writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
            _writer.WriteAttributeString("Target", link.Url);
            _writer.WriteAttributeString("TargetMode", "External");
            _writer.WriteEndElement();
        }
    }
}

/// <summary>
/// Holds computed relationship IDs for all document parts. 0 means not present.
/// </summary>
public class DocumentRelationshipIds
{
    public int SettingsRId { get; set; }
    public int FontTableRId { get; set; }
    public int ThemeRId { get; set; }
    public int FirstImageRId { get; set; }
    public int FirstChartRId { get; set; }
    public int FirstOleRId { get; set; }
    public int NumberingRId { get; set; }
    public int FootnotesRId { get; set; }
    public int EndnotesRId { get; set; }
    public int VbaProjectRId { get; set; }
    // Aggregate header/footer IDs (kept for backward compatibility)
    public int HeaderRId { get; set; }
    public int FooterRId { get; set; }
    // Per-type header IDs
    public int HeaderFirstRId { get; set; }
    public int HeaderOddRId { get; set; }
    public int HeaderEvenRId { get; set; }
    // Per-type footer IDs
    public int FooterFirstRId { get; set; }
    public int FooterOddRId { get; set; }
    public int FooterEvenRId { get; set; }
    /// <summary>Highest relationship ID reserved for non-hyperlink relationships.</summary>
    public int LastUsedRId { get; set; }
}

/// <summary>
/// Settings XML Writer
/// </summary>
public class SettingsWriter
{
    private readonly XmlWriter _writer;
    
    public SettingsWriter(XmlWriter writer)
    {
        _writer = writer;
    }
    
    /// <summary>
    /// Writes the settings part
    /// </summary>
    public void WriteSettings()
    {
        _writer.WriteStartDocument();
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "settings", wNs);
        
        // Force update fields on open (essential for TOC and page numbers)
        _writer.WriteStartElement("w", "updateFields", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "true");
        _writer.WriteEndElement();

        WriteZoom();
        WriteProofState();
        WriteDefaultTabStop();
        WriteHyphenationZone();
        WriteCharacterSpacing();
        
        _writer.WriteEndElement(); // w:settings
        _writer.WriteEndDocument();
    }
    
    private void WriteZoom()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "zoom", wNs);
        _writer.WriteAttributeString("w", "percent", wNs, "100");
        _writer.WriteEndElement();
    }
    
    private void WriteProofState()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "proofState", wNs);
        _writer.WriteAttributeString("w", "spelling", wNs, "clean");
        _writer.WriteAttributeString("w", "grammar", wNs, "clean");
        _writer.WriteEndElement();
    }
    
    private void WriteDefaultTabStop()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "defaultTabStop", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "720");
        _writer.WriteEndElement();
    }
    
    private void WriteHyphenationZone()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "hyphenationZone", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "360");
        _writer.WriteEndElement();
    }
    
    private void WriteCharacterSpacing()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "characterSpacingControl", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "doNotCompress");
        _writer.WriteEndElement();
    }
}

/// <summary>
/// App Properties Writer
/// </summary>
public class AppPropertiesWriter
{
    private readonly XmlWriter _writer;
    
    public AppPropertiesWriter(XmlWriter writer)
    {
        _writer = writer;
    }
    
    /// <summary>
    /// Writes the app properties
    /// </summary>
    public void WriteAppProperties()
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        
        WriteElement("Template", "Normal.dotm");
        WriteElement("Manager", "");
        WriteElement("Company", "");
        WriteElement("Application", "Nedev.DocToDocx");
        WriteElement("AppVersion", "15.0000");
        
        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }
    
    private void WriteElement(string name, string value)
    {
        if (name.Contains(':'))
        {
            var parts = name.Split(':');
            _writer.WriteStartElement(parts[0], parts[1], GetNamespaceForPrefix(parts[0]));
            _writer.WriteString(value);
            _writer.WriteEndElement();
        }
        else
        {
            _writer.WriteElementString(name, value);
        }
    }

    private string GetNamespaceForPrefix(string prefix)
    {
        return prefix switch
        {
            "dc" => "http://purl.org/dc/elements/1.1/",
            "cp" => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dcterms" => "http://purl.org/dc/terms/",
            "xsi" => "http://www.w3.org/2001/XMLSchema-instance",
            _ => ""
        };
    }
}

/// <summary>
/// Core Properties Writer (metadata)
/// </summary>
public class CorePropertiesWriter
{
    private readonly XmlWriter _writer;
    private readonly DocumentProperties _props;
    
    public CorePropertiesWriter(XmlWriter writer, DocumentProperties props)
    {
        _writer = writer;
        _props = props;
    }
    
    /// <summary>
    /// Writes the core properties
    /// </summary>
    public void WriteCoreProperties()
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        
        // Add XML namespace definitions
        _writer.WriteAttributeString("xmlns", "cp", null, "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        _writer.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
        _writer.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
        _writer.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
        _writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
        
        WriteElement("dc:title", _props.Title ?? "");
        WriteElement("dc:subject", _props.Subject ?? "");
        WriteElement("dc:creator", _props.Author ?? "Nedev.DocToDocx");
        WriteElement("cp:keywords", _props.Keywords ?? "");
        WriteElement("dc:description", _props.Comments ?? "");
        WriteElement("cp:lastModifiedBy", _props.Author ?? "Nedev.DocToDocx");
        
        _writer.WriteStartElement("dcterms", "created", "http://purl.org/dc/terms/");
        _writer.WriteAttributeString("xsi", "type", "http://www.w3.org/2001/XMLSchema-instance", "dcterms:W3CDTF");
        var created = _props.Created > new DateTime(1900, 1, 1) ? _props.Created : DateTime.UtcNow;
        _writer.WriteString(created.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("dcterms", "modified", "http://purl.org/dc/terms/");
        _writer.WriteAttributeString("xsi", "type", "http://www.w3.org/2001/XMLSchema-instance", "dcterms:W3CDTF");
        var modified = _props.Modified > new DateTime(1900, 1, 1) ? _props.Modified : DateTime.UtcNow;
        _writer.WriteString(modified.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        _writer.WriteEndElement();
        
        WriteElement("cp:category", "");
        WriteElement("cp:contentStatus", "Draft");
        
        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }
    
    private void WriteElement(string name, string value)
    {
        if (name.Contains(':'))
        {
            var parts = name.Split(':');
            _writer.WriteStartElement(parts[0], parts[1], GetNamespaceForPrefix(parts[0]));
            _writer.WriteString(value);
            _writer.WriteEndElement();
        }
        else
        {
            _writer.WriteElementString(name, value);
        }
    }

    private string GetNamespaceForPrefix(string prefix)
    {
        return prefix switch
        {
            "dc" => "http://purl.org/dc/elements/1.1/",
            "cp" => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dcterms" => "http://purl.org/dc/terms/",
            "xsi" => "http://www.w3.org/2001/XMLSchema-instance",
            _ => ""
        };
    }
}

/// <summary>
/// Content Types XML Writer
/// </summary>
public class ContentTypesWriter
{
    private readonly XmlWriter _writer;
    
    public ContentTypesWriter(XmlWriter writer)
    {
        _writer = writer;
    }
    
    /// <summary>
    /// Writes the [Content_Types].xml
    /// </summary>
    public void WriteContentTypes(DocumentModel document)
    {
        _writer.WriteStartDocument();
        
        // Write types root
        _writer.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
        _writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
        
        // Default types
        WriteDefault("xml", "application/xml");
        WriteDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
        WriteDefault("bin", "application/vnd.openxmlformats-officedocument.oleObject");
        WriteDefault("png", "image/png");
        WriteDefault("jpg", "image/jpeg");
        WriteDefault("jpeg", "image/jpeg");
        WriteDefault("gif", "image/gif");
        WriteDefault("bmp", "image/bmp");
        WriteDefault("odttf", "application/vnd.openxmlformats-officedocument.obfuscatedFont");
        WriteDefault("ttf", "application/x-font-ttf");
        
        // Override types
        WriteOverride("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
        WriteOverride("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");

        // Main document type changes if macro-enabled
        string mainType = document.VbaProject != null 
            ? "application/vnd.ms-word.document.macroEnabled.main+xml" 
            : "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        WriteOverride("/word/document.xml", mainType);
        
        WriteOverride("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
        
        // Font Table
        if (document.Styles.Fonts.Any(f => f.EmbeddedData != null))
        {
            WriteOverride("/word/fontTable.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml");
        }
        
        // Theme
        if (!string.IsNullOrEmpty(document.Theme.XmlContent))
        {
            WriteOverride("/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml");
        }
        
        // VbaProject
        if (document.VbaProject != null)
        {
            WriteOverride("/word/vbaProject.bin", "application/vnd.ms-office.vbaProject");
        }
        
        
        // Settings (always present)
        WriteOverride("/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");

        
        // Headers and footers: register up to three header/footer parts
        bool hasHeaderFirst = document.HeadersFooters.Headers.Any(h => h.Type == HeaderFooterType.HeaderFirst);
        bool hasHeaderOdd = document.HeadersFooters.Headers.Any(h => h.Type == HeaderFooterType.HeaderOdd);
        bool hasHeaderEven = document.HeadersFooters.Headers.Any(h => h.Type == HeaderFooterType.HeaderEven);

        if (hasHeaderFirst)
        {
            WriteOverride("/word/header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
        }
        if (hasHeaderOdd)
        {
            WriteOverride("/word/header2.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
        }
        if (hasHeaderEven)
        {
            WriteOverride("/word/header3.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
        }
        
        bool hasFooterFirst = document.HeadersFooters.Footers.Any(f => f.Type == HeaderFooterType.FooterFirst);
        bool hasFooterOdd = document.HeadersFooters.Footers.Any(f => f.Type == HeaderFooterType.FooterOdd);
        bool hasFooterEven = document.HeadersFooters.Footers.Any(f => f.Type == HeaderFooterType.FooterEven);

        if (hasFooterFirst)
        {
            WriteOverride("/word/footer1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
        }
        if (hasFooterOdd)
        {
            WriteOverride("/word/footer2.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
        }
        if (hasFooterEven)
        {
            WriteOverride("/word/footer3.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
        }
        
        // Footnotes and endnotes
        if (document.Footnotes.Count > 0)
        {
            WriteOverride("/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml");
        }
        
        if (document.Endnotes.Count > 0)
        {
            WriteOverride("/word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml");
        }
        
        // Numbering
        if (document.Paragraphs.Any(p => p.ListFormatId > 0) || document.NumberingDefinitions.Count > 0)
        {
            WriteOverride("/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
        }

        // Charts
        if (document.Charts.Count > 0)
        {
            for (int i = 0; i < document.Charts.Count; i++)
            {
                var partName = $"/word/charts/chart{i + 1}.xml";
                WriteOverride(partName, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
            }
        }
        
        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }
    
    private void WriteDefault(string extension, string contentType)
    {
        _writer.WriteStartElement("Default");
        _writer.WriteAttributeString("Extension", extension);
        _writer.WriteAttributeString("ContentType", contentType);
        _writer.WriteEndElement();
    }
    
    private void WriteOverride(string partName, string contentType)
    {
        _writer.WriteStartElement("Override");
        _writer.WriteAttributeString("PartName", partName);
        _writer.WriteAttributeString("ContentType", contentType);
        _writer.WriteEndElement();
    }
}
