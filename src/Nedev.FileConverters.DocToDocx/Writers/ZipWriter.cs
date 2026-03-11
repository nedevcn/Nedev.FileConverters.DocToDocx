using System.IO.Compression;
using System.Text;
using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;
using Nedev.FileConverters.DocToDocx.Writers;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// ZIP package writer - creates DOCX file without third-party dependencies
/// Uses System.IO.Compression for ZIP support
/// </summary>
public class ZipWriter : IDisposable
{
    private readonly Stream _outputStream;
    private readonly ZipArchive _archive;
    
    public ZipWriter(Stream outputStream)
    {
        _outputStream = outputStream;
        _archive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);
    }
    
    /// <summary>
    /// Adds an XML entry to the archive
    /// </summary>
    public void AddXmlEntry(string entryName, Action<XmlWriter> writeAction)
    {
        var entry = _archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = entry.Open();
        
        var settings = new XmlWriterSettings
        {
            Indent = false,
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false,
            NewLineHandling = NewLineHandling.None
        };
        
        using var streamWriter = new StreamWriter(stream, new UTF8Encoding(false));
        using var xmlWriter = XmlWriter.Create(streamWriter, settings);
        writeAction(xmlWriter);
        xmlWriter.Flush();
        streamWriter.Flush();
    }
    
    /// <summary>
    /// Adds a binary entry to the archive
    /// </summary>
    public void AddBinaryEntry(string entryName, byte[] data)
    {
        var entry = _archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = entry.Open();
        stream.Write(data, 0, data.Length);
    }
    
    /// <summary>
    /// Adds a binary entry to the archive from a stream
    /// </summary>
    public void AddBinaryEntry(string entryName, Stream dataStream, int length)
    {
        var entry = _archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = entry.Open();
        
        var buffer = new byte[8192];
        var totalRead = 0;
        
        while (totalRead < length)
        {
            var read = dataStream.Read(buffer, 0, Math.Min(buffer.Length, length - totalRead));
            if (read == 0) break;
            stream.Write(buffer, 0, read);
            totalRead += read;
        }
    }
    
    /// <summary>
    /// Writes the document to the archive using default options.
    /// </summary>
    public void WriteDocument(DocumentModel document)
    {
        WriteDocument(document, options: null);
    }

    /// <summary>
    /// Writes the document to the archive with the provided writer options.
    /// </summary>
    public void WriteDocument(DocumentModel document, DocumentWriterOptions? options)
    {
        // Assign relationship IDs for hyperlinks after other IDs are reserved
        var relIds = RelationshipsWriter.ComputeRelationshipIds(document);
        HyperlinkIdAssigner.AssignHyperlinkIds(document, relIds.LastUsedRId + 1);

        // Write [Content_Types].xml
        AddXmlEntry("[Content_Types].xml", w =>
        {
            var writer = new ContentTypesWriter(w);
            writer.WriteContentTypes(document);
        });
        
        // Write _rels/.rels
        AddXmlEntry("_rels/.rels", w =>
        {
            var writer = new RelationshipsWriter(w);
            writer.WriteMainRelationships();
        });
        
        // Write docProps/core.xml
        AddXmlEntry("docProps/core.xml", w =>
        {
            var writer = new CorePropertiesWriter(w, document.Properties);
            writer.WriteCoreProperties();
        });
        
        // Write docProps/app.xml
        AddXmlEntry("docProps/app.xml", w =>
        {
            var writer = new AppPropertiesWriter(w);
            writer.WriteAppProperties();
        });
        
        // Write word/document.xml
        AddXmlEntry("word/document.xml", w =>
        {
            var writer = new DocumentWriter(w, options);
            writer.WriteDocument(document);
        });

        // Write word/theme/theme1.xml if present
        if (!string.IsNullOrEmpty(document.Theme.XmlContent))
        {
            AddBinaryEntry("word/theme/theme1.xml", Encoding.UTF8.GetBytes(document.Theme.XmlContent));
        }
        
        // Write word/_rels/document.xml.rels
        AddXmlEntry("word/_rels/document.xml.rels", w =>
        {
            var writer = new RelationshipsWriter(w);
            writer.WriteDocumentRelationships(document, includeHyperlinks: options?.EnableHyperlinks ?? true);
        });
        
        // Write word/styles.xml
        AddXmlEntry("word/styles.xml", w =>
        {
            var writer = new StylesWriter(w);
            writer.WriteStyles(document);
        });
        
        // Write word/fontTable.xml and font files
        var embeddedFonts = document.Styles.Fonts.Where(f => f.EmbeddedData != null && f.EmbeddedData.Length > 0).ToList();
        if (embeddedFonts.Count > 0)
        {
            // Dictionary to map font name to relationship ID for fontTable.xml
            var fontRelIds = new Dictionary<string, string>();
            var fontRels = new List<(string rId, string Type, string Target)>();
            
            for (int i = 0; i < embeddedFonts.Count; i++)
            {
                var font = embeddedFonts[i];
                var rId = $"rId{i + 1}";
                fontRelIds[font.Name] = rId;
                
                // Use .odttf and obfuscate if you want standard Word behavior
                // Let's use obfuscation as it is safer for compatibility
                var fontKey = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
                var obfuscatedData = FontObfuscator.ObfuscateFont(font.EmbeddedData, fontKey);
                
                string fontFilePath = $"fonts/font{i + 1}.odttf";
                fontRels.Add((rId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font", fontFilePath));
                
                AddBinaryEntry($"word/{fontFilePath}", obfuscatedData);
            }
            
            // Write word/_rels/fontTable.xml.rels
            AddXmlEntry("word/_rels/fontTable.xml.rels", w => WritePartRels(w, fontRels));

            // Write word/fontTable.xml
            AddXmlEntry("word/fontTable.xml", w =>
            {
                var writer = new FontTableWriter(w);
                writer.WriteFontTable(document, fontRelIds);
            });
        }
        
        // Write word/numbering.xml if document has lists
        if (document.Paragraphs.Any(p => p.ListFormatId > 0) || document.NumberingDefinitions.Count > 0)
        {
            AddXmlEntry("word/numbering.xml", w =>
            {
                var writer = new NumberingWriter(w);
                writer.WriteNumbering(document);
            });
        }

        // Write word/vbaProject.bin
        if (document.VbaProject != null)
        {
            AddBinaryEntry("word/vbaProject.bin", document.VbaProject);
        }
        
        // Write word/settings.xml (always present)
        AddXmlEntry("word/settings.xml", w =>
        {
            var writer = new SettingsWriter(w);
            writer.WriteSettings(document);
        });
        
        // Write images (use minimal 1x1 PNG when image has no data to avoid broken/corrupt part)
        for (int i = 0; i < document.Images.Count; i++)
        {
            var image = document.Images[i];
            var extension = GetImageExtension(image.Type);
            var data = image.Data;
            if (data == null || data.Length == 0)
            {
                data = MinimalTransparentPng;
                extension = ".png";
            }
            AddBinaryEntry($"word/media/image{i + 1}{extension}", data);
        }
        
        // Write headers and footers - simplified approach: create only one header and one footer file
        if (document.HeadersFooters.Headers.Count > 0 || document.HeadersFooters.Footers.Count > 0)
        {
            WriteHeadersAndFooters(document);
        }
        
        // Write footnotes
        if (document.Footnotes.Count > 0)
        {
            WriteFootnotes(document);
        }
        
        // Write endnotes
        if (document.Endnotes.Count > 0)
        {
            WriteEndnotes(document);
        }

        // Write comments/annotations
        if (document.Annotations != null && document.Annotations.Count > 0)
        {
            AddXmlEntry("word/comments.xml", w =>
            {
                var writer = new CommentsWriter(w);
                writer.WriteComments(document);
            });
        }
        
        // Write charts (if any). For now we emit one chart part per ChartModel
        // using a very small, self-contained ChartsWriter.
        if (document.Charts.Count > 0)
        {
            WriteCharts(document);
        }
        
        // Write OLE Objects
        if (document.OleObjects.Count > 0)
        {
            for (int i = 0; i < document.OleObjects.Count; i++)
            {
                var ole = document.OleObjects[i];
                if (ole.ObjectData != null && ole.ObjectData.Length > 0)
                {
                    AddBinaryEntry($"word/embeddings/oleObject{i + 1}.bin", ole.ObjectData);
                }
            }
        }
    }
    
    /// <summary>
    /// Disposes the writer
    /// </summary>
    public void Dispose()
    {
        _archive?.Dispose();
    }
    
    /// <summary>Minimal 1x1 transparent PNG (67 bytes) for placeholder when image has no data.</summary>
    private static readonly byte[] MinimalTransparentPng = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVQYV2NgYAAAAAMAAWgmWQ0AAAAASUVORK5CYII=");

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
    /// Writes headers and footers to the archive.
    /// </summary>
    private void WriteHeadersAndFooters(DocumentModel document)
    {
        foreach (var header in RelationshipsWriter.GetUsableHeaderParts(document))
        {
            var (partRels, imageMap, oleMap) = BuildHeaderFooterPartRels(document, header);
            if (partRels.Count > 0 && !string.IsNullOrEmpty(header.PartName))
            {
                AddXmlEntry($"word/_rels/{header.PartName}.rels", w => WritePartRels(w, partRels));
            }

            if (string.IsNullOrEmpty(header.PartName))
                continue;

            AddXmlEntry($"word/{header.PartName}", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteHeader(header, document, imageMap, oleMap);
            });
        }

        foreach (var footer in RelationshipsWriter.GetUsableFooterParts(document))
        {
            var (partRels, imageMap, oleMap) = BuildHeaderFooterPartRels(document, footer);
            if (partRels.Count > 0 && !string.IsNullOrEmpty(footer.PartName))
            {
                AddXmlEntry($"word/_rels/{footer.PartName}.rels", w => WritePartRels(w, partRels));
            }

            if (string.IsNullOrEmpty(footer.PartName))
                continue;

            AddXmlEntry($"word/{footer.PartName}", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteFooter(footer, document, imageMap, oleMap);
            });
        }
    }
    
    /// <summary>
    /// Writes footnotes to the archive (and footnote part relationships if images are used).
    /// </summary>
    private void WriteFootnotes(DocumentModel document)
    {
        var (noteRels, imageMap) = BuildNotePartImageRels(document, document.Footnotes.Cast<NoteModelBase>().ToList());
        if (noteRels.Count > 0)
        {
            AddXmlEntry("word/_rels/footnotes.xml.rels", w => WritePartRels(w, noteRels));
        }
        AddXmlEntry("word/footnotes.xml", w =>
        {
            var writer = new FootnotesWriter(w);
            writer.WriteFootnotes(document.Footnotes, document, imageMap);
        });
    }

    /// <summary>
    /// Writes endnotes to the archive (and endnote part relationships if images are used).
    /// </summary>
    private void WriteEndnotes(DocumentModel document)
    {
        var (noteRels, imageMap) = BuildNotePartImageRels(document, document.Endnotes.Cast<NoteModelBase>().ToList());
        if (noteRels.Count > 0)
        {
            AddXmlEntry("word/_rels/endnotes.xml.rels", w => WritePartRels(w, noteRels));
        }
        AddXmlEntry("word/endnotes.xml", w =>
        {
            var writer = new FootnotesWriter(w);
            writer.WriteEndnotes(document.Endnotes, document, imageMap);
        });
    }

    /// <summary>
    /// Collects image indices used in notes and builds (rels entries, imageIndex -> rId map).
    /// </summary>
    private static (List<(string rId, string Type, string Target)> rels, Dictionary<int, string> imageIndexToRelId) BuildNotePartImageRels(DocumentModel document, List<NoteModelBase> notes)
    {
        var order = new List<int>();
        var seen = new HashSet<int>();
        foreach (var note in notes)
        {
            foreach (var para in note.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    if (!run.IsPicture || run.ImageIndex < 0 || run.ImageIndex >= document.Images.Count) continue;
                    if (seen.Add(run.ImageIndex))
                        order.Add(run.ImageIndex);
                }
            }
        }
        var rels = new List<(string rId, string Type, string Target)>();
        var imageIndexToRelId = new Dictionary<int, string>();
        for (int i = 0; i < order.Count; i++)
        {
            var imageIndex = order[i];
            var ext = GetImageExtension(document.Images[imageIndex].Type);
            var rId = $"rId{i + 1}";
            rels.Add((rId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", $"../media/image{imageIndex + 1}{ext}"));
            imageIndexToRelId[imageIndex] = rId;
        }
        return (rels, imageIndexToRelId);
    }

    private static (List<(string rId, string Type, string Target)> rels, Dictionary<int, string> imageIndexToRelId, Dictionary<string, string> oleObjectIdToRelId) BuildHeaderFooterPartRels(DocumentModel document, HeaderFooterModel headerFooter)
    {
        var rels = new List<(string rId, string Type, string Target)>();
        var imageIndexToRelId = new Dictionary<int, string>();
        var oleObjectIdToRelId = new Dictionary<string, string>(StringComparer.Ordinal);
        var nextId = 1;

        if (headerFooter.Paragraphs == null)
            return (rels, imageIndexToRelId, oleObjectIdToRelId);

        foreach (var paragraph in headerFooter.Paragraphs)
        {
            foreach (var run in paragraph.Runs)
            {
                if (run.ImageIndex >= 0 && run.ImageIndex < document.Images.Count && !imageIndexToRelId.ContainsKey(run.ImageIndex))
                {
                    var ext = GetImageExtension(document.Images[run.ImageIndex].Type);
                    var rId = $"rId{nextId++}";
                    rels.Add((rId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", $"../media/image{run.ImageIndex + 1}{ext}"));
                    imageIndexToRelId[run.ImageIndex] = rId;
                }

                if (run.IsOle && !string.IsNullOrEmpty(run.OleObjectId) && !oleObjectIdToRelId.ContainsKey(run.OleObjectId))
                {
                    var oleIndex = document.OleObjects.FindIndex(o => o.ObjectId == run.OleObjectId);
                    if (oleIndex < 0)
                        continue;

                    var rId = $"rId{nextId++}";
                    rels.Add((rId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", $"../embeddings/oleObject{oleIndex + 1}.bin"));
                    oleObjectIdToRelId[run.OleObjectId] = rId;
                }
            }
        }

        return (rels, imageIndexToRelId, oleObjectIdToRelId);
    }

    private static void WritePartRels(XmlWriter w, List<(string rId, string Type, string Target)> rels)
    {
        w.WriteStartDocument();
        w.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (var (rId, type, target) in rels)
        {
            w.WriteStartElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships");
            w.WriteAttributeString("Id", rId);
            w.WriteAttributeString("Type", type);
            w.WriteAttributeString("Target", target);
            w.WriteEndElement();
        }
        w.WriteEndElement();
        w.WriteEndDocument();
    }

    /// <summary>
    /// Writes chart parts (word/charts/chartN.xml) for all charts in the document.
    /// </summary>
    private void WriteCharts(DocumentModel document)
    {
        for (int i = 0; i < document.Charts.Count; i++)
        {
            var chart = document.Charts[i];
            chart.Index = i;
            AddXmlEntry($"word/charts/chart{i + 1}.xml", w =>
            {
                var writer = new ChartsWriter(w);
                writer.WriteChart(chart);
            });
        }
    }
}



