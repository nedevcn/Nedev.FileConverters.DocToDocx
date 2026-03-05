using System.IO.Compression;
using System.Text;
using System.Xml;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Writers;

namespace Nedev.DocToDocx.Writers;

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
    /// Writes the document to the archive
    /// </summary>
    public void WriteDocument(DocumentModel document)
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
            var writer = new DocumentWriter(w);
            writer.WriteDocument(document);
        });
        
        // Write word/_rels/document.xml.rels
        AddXmlEntry("word/_rels/document.xml.rels", w =>
        {
            var writer = new RelationshipsWriter(w);
            writer.WriteDocumentRelationships(document);
        });
        
        // Write word/styles.xml
        AddXmlEntry("word/styles.xml", w =>
        {
            var writer = new StylesWriter(w);
            writer.WriteStyles(document);
        });
        
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
            writer.WriteSettings();
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
    /// Generates up to three header parts (header1/2/3.xml) and three footer
    /// parts (footer1/2/3.xml) corresponding to First/Odd/Even.
    /// </summary>
    private void WriteHeadersAndFooters(DocumentModel document)
    {
        var headers = document.HeadersFooters.Headers;
        var footers = document.HeadersFooters.Footers;

        // HeaderFirst -> header1.xml
        var headerFirst = headers.FirstOrDefault(h => h.Type == HeaderFooterType.HeaderFirst);
        if (headerFirst != null)
        {
            AddXmlEntry("word/header1.xml", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteHeader(headerFirst, document);
            });
        }

        // HeaderOdd (default) -> header2.xml
        var headerOdd = headers.FirstOrDefault(h => h.Type == HeaderFooterType.HeaderOdd);
        if (headerOdd != null)
        {
            AddXmlEntry("word/header2.xml", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteHeader(headerOdd, document);
            });
        }

        // HeaderEven -> header3.xml
        var headerEven = headers.FirstOrDefault(h => h.Type == HeaderFooterType.HeaderEven);
        if (headerEven != null)
        {
            AddXmlEntry("word/header3.xml", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteHeader(headerEven, document);
            });
        }

        // FooterFirst -> footer1.xml
        var footerFirst = footers.FirstOrDefault(f => f.Type == HeaderFooterType.FooterFirst);
        if (footerFirst != null)
        {
            AddXmlEntry("word/footer1.xml", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteFooter(footerFirst, document);
            });
        }

        // FooterOdd (default) -> footer2.xml
        var footerOdd = footers.FirstOrDefault(f => f.Type == HeaderFooterType.FooterOdd);
        if (footerOdd != null)
        {
            AddXmlEntry("word/footer2.xml", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteFooter(footerOdd, document);
            });
        }

        // FooterEven -> footer3.xml
        var footerEven = footers.FirstOrDefault(f => f.Type == HeaderFooterType.FooterEven);
        if (footerEven != null)
        {
            AddXmlEntry("word/footer3.xml", w =>
            {
                var writer = new HeaderFooterWriter(w);
                writer.WriteFooter(footerEven, document);
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
    private static (List<(string rId, string Target)> rels, Dictionary<int, string> imageIndexToRelId) BuildNotePartImageRels(DocumentModel document, List<NoteModelBase> notes)
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
        var rels = new List<(string rId, string Target)>();
        var imageIndexToRelId = new Dictionary<int, string>();
        for (int i = 0; i < order.Count; i++)
        {
            var imageIndex = order[i];
            var ext = GetImageExtension(document.Images[imageIndex].Type);
            var rId = $"rId{i + 1}";
            rels.Add((rId, $"../media/image{imageIndex + 1}{ext}"));
            imageIndexToRelId[imageIndex] = rId;
        }
        return (rels, imageIndexToRelId);
    }

    private static void WritePartRels(XmlWriter w, List<(string rId, string Target)> rels)
    {
        w.WriteStartDocument();
        w.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (var (rId, target) in rels)
        {
            w.WriteStartElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships");
            w.WriteAttributeString("Id", rId);
            w.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
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

/// <summary>
/// Simple ZIP entry structure for manual ZIP creation
/// </summary>
internal class ZipEntry
{
    public string Name { get; set; } = "";
    public long CompressedSize { get; set; }
    public long UncompressedSize { get; set; }
    public ushort CompressionMethod { get; set; }
    public uint Crc32 { get; set; }
    public long LocalHeaderOffset { get; set; }
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public byte[] CompressedData { get; set; } = Array.Empty<byte>();
}

/// <summary>
/// Manual ZIP writer without System.IO.Compression (alternative implementation)
/// Uses DeflateStream directly for maximum control
/// </summary>
public class ManualZipWriter : IDisposable
{
    private readonly Stream _outputStream;
    private readonly List<LocalFileHeader> _entries = new();
    
    public ManualZipWriter(string outputPath)
    {
        _outputStream = File.Create(outputPath);
    }
    
    public ManualZipWriter(Stream outputStream)
    {
        _outputStream = outputStream;
    }
    
    /// <summary>
    /// Creates a new entry in the ZIP file
    /// </summary>
    public void CreateEntry(string name, byte[] data, bool compress = true)
    {
        var entry = new LocalFileHeader
        {
            FileName = name,
            UncompressedSize = data.Length,
            CompressedData = compress ? CompressDeflate(data) : data,
            CompressionMethod = compress ? (ushort)8 : (ushort)0,
            Crc32 = CalculateCrc32(data)
        };

        // Set compressed size based on actual compressed buffer length
        entry.CompressedSize = entry.CompressedData.Length;
        
        _entries.Add(entry);
    }
    
    /// <summary>
    /// Writes all entries to the output stream
    /// </summary>
    public void Finish()
    {
        // Write local file headers
        foreach (var entry in _entries)
        {
            entry.LocalHeaderOffset = _outputStream.Position;
            WriteLocalFileHeader(entry);
        }
        
        // Write central directory
        var centralDirectoryOffset = _outputStream.Position;
        WriteCentralDirectory();
        
        // Write end of central directory
        WriteEndOfCentralDirectory(centralDirectoryOffset);
    }
    
    private void WriteLocalFileHeader(LocalFileHeader entry)
    {
        // Local file header signature
        _outputStream.WriteByte(0x50);
        _outputStream.WriteByte(0x4B);
        _outputStream.WriteByte(0x03);
        _outputStream.WriteByte(0x04);
        
        // Version needed to extract
        WriteUInt16LE(20);
        
        // General purpose bit flag
        WriteUInt16LE(0);
        
        // Compression method
        WriteUInt16LE(entry.CompressionMethod);
        
        // File last modification time
        WriteUInt16LE(0);
        
        // File last modification date
        WriteUInt16LE(0);
        
        // CRC-32
        WriteUInt32LE(entry.Crc32);
        
        // Compressed size
        WriteUInt32LE((uint)entry.CompressedSize);
        
        // Uncompressed size
        WriteUInt32LE((uint)entry.UncompressedSize);
        
        // File name length
        var fileNameBytes = Encoding.UTF8.GetBytes(entry.FileName);
        WriteUInt16LE((ushort)fileNameBytes.Length);
        
        // Extra field length
        WriteUInt16LE(0);
        
        // File name
        _outputStream.Write(fileNameBytes);
        
        // Compressed data
        _outputStream.Write(entry.CompressedData);
    }
    
    private void WriteCentralDirectory()
    {
        foreach (var entry in _entries)
        {
            // Central file header signature
            _outputStream.WriteByte(0x50);
            _outputStream.WriteByte(0x4B);
            _outputStream.WriteByte(0x01);
            _outputStream.WriteByte(0x02);
            
            // Version made by
            WriteUInt16LE(20);
            
            // Version needed to extract
            WriteUInt16LE(20);
            
            // General purpose bit flag
            WriteUInt16LE(0);
            
            // Compression method
            WriteUInt16LE(entry.CompressionMethod);
            
            // File last modification time
            WriteUInt16LE(0);
            
            // File last modification date
            WriteUInt16LE(0);
            
            // CRC-32
            WriteUInt32LE(entry.Crc32);
            
            // Compressed size
            WriteUInt32LE((uint)entry.CompressedSize);
            
            // Uncompressed size
            WriteUInt32LE((uint)entry.UncompressedSize);
            
            // File name length
            var fileNameBytes = Encoding.UTF8.GetBytes(entry.FileName);
            WriteUInt16LE((ushort)fileNameBytes.Length);
            
            // Extra field length
            WriteUInt16LE(0);
            
            // File comment length
            WriteUInt16LE(0);
            
            // Disk number start
            WriteUInt16LE(0);
            
            // Internal file attributes
            WriteUInt16LE(0);
            
            // External file attributes
            WriteUInt32LE(0);
            
            // Relative offset of local header
            WriteUInt32LE((uint)entry.LocalHeaderOffset);
            
            // File name
            _outputStream.Write(fileNameBytes);
        }
    }
    
    private void WriteEndOfCentralDirectory(long centralDirectoryOffset)
    {
        // End of central dir signature
        _outputStream.WriteByte(0x50);
        _outputStream.WriteByte(0x4B);
        _outputStream.WriteByte(0x05);
        _outputStream.WriteByte(0x06);
        
        // Number of this disk
        WriteUInt16LE(0);
        
        // Number of the disk with the start of the central directory
        WriteUInt16LE(0);
        
        // Total number of entries in the central directory on this disk
        WriteUInt16LE((ushort)_entries.Count);
        
        // Total number of entries in the central directory
        WriteUInt16LE((ushort)_entries.Count);
        
        // Size of the central directory
        var centralDirectorySize = _outputStream.Position - centralDirectoryOffset;
        WriteUInt32LE((uint)centralDirectorySize);
        
        // Offset of start of central directory
        WriteUInt32LE((uint)centralDirectoryOffset);
        
        // .ZIP file comment length
        WriteUInt16LE(0);
    }
    
    private void WriteUInt16LE(ushort value)
    {
        var bytes = BitConverter.GetBytes(value);
        _outputStream.Write(bytes);
    }
    
    private void WriteUInt32LE(uint value)
    {
        var bytes = BitConverter.GetBytes(value);
        _outputStream.Write(bytes);
    }
    
    private byte[] CompressDeflate(byte[] data)
    {
        using var output = new MemoryStream();
        using var deflate = new DeflateStream(output, CompressionLevel.Optimal);
        deflate.Write(data, 0, data.Length);
        deflate.Flush();
        return output.ToArray();
    }
    
    private uint CalculateCrc32(byte[] data)
    {
        // Simple CRC-32 implementation
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
            {
                if ((crc & 1) != 0)
                    crc = (crc >> 1) ^ 0xEDB88320;
                else
                    crc >>= 1;
            }
        }
        return ~crc;
    }
    
    public void Dispose()
    {
        _outputStream?.Dispose();
    }
}

internal class LocalFileHeader
{
    public string FileName { get; set; } = "";
    public long UncompressedSize { get; set; }
    public long CompressedSize { get; set; }
    public ushort CompressionMethod { get; set; }
    public uint Crc32 { get; set; }
    public long LocalHeaderOffset { get; set; }
    public byte[] CompressedData { get; set; } = Array.Empty<byte>();
}
