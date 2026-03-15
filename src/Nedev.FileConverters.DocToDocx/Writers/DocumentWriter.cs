using System.Text.RegularExpressions;
using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// Options that influence how a <see cref="DocumentWriter"/> emits XML.
/// </summary>
public class DocumentWriterOptions
{
    /// <summary>
    /// When <c>true</c> (the default) hyperlink runs produce
    /// <c>&lt;w:hyperlink&gt;</c> elements and external relationships.
    /// When <c>false</c> the text is written as plain runs and no
    /// hyperlink relationship is created, which avoids the Word prompt
    /// about linked fields.  Clients can disable hyperlinks if they
    /// prefer a warning-free, static document.
    /// </summary>
    public bool EnableHyperlinks { get; set; } = true;
}

/// <summary>
/// Writes DOCX document using XmlWriter for optimal streaming performance
/// </summary>
public partial class DocumentWriter
{
    private readonly XmlWriter _writer;
    private readonly DocumentWriterOptions _options;
    private int _runId = 0;
    private int _trackChangeId = 1;
    private DocumentModel? _document;
    private DocumentRelationshipIds? _relationshipIds;
    private readonly Dictionary<string, int> _bookmarkIds = new(StringComparer.Ordinal);
    private int _bookmarkCounter = 0;
    private HashSet<string> _startedComments = new();
    private HashSet<string> _endedComments = new();
    /// <summary>Paragraph index → list of annotation IDs whose range starts at that paragraph.</summary>
    private Dictionary<int, List<int>> _commentStartsByParagraph = new();
    /// <summary>Paragraph index → list of annotation IDs whose range ends at that paragraph.</summary>
    private Dictionary<int, List<int>> _commentEndsByParagraph = new();
    private IReadOnlyDictionary<int, string>? _imageRelationshipOverrides;
    private IReadOnlyDictionary<string, string>? _oleRelationshipOverrides;
    /// <summary>When true, do not emit pageBreakBefore so leading content (e.g. 绿色等级评价报告) stays on page 1.</summary>
    private bool _suppressLeadingPageBreak;

    // Compiled regex patterns for better performance
    private static readonly Regex HyperlinkFieldRegex = new(
        "\\s*HYPERLINK\\s+\"[^\"]*\"\\s*",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex MultipleSpacesRegex = new(
        "[ ]{2,}",
        RegexOptions.Compiled);

    /// <summary>
    /// Creates a document writer for the target XML stream.
    /// </summary>
    /// <param name="writer">The XML writer that receives the document markup.</param>
    /// <param name="options">Optional writer behavior flags.</param>
    public DocumentWriter(XmlWriter writer, DocumentWriterOptions? options = null)
    {
        _writer = writer;
        _options = options ?? new DocumentWriterOptions();
    }

    /// <summary>
    /// Binds document-scoped context for fragment writers that reuse paragraph/run emission.
    /// </summary>
    internal DocumentWriter BindDocumentContext(
        DocumentModel document,
        DocumentRelationshipIds? relationshipIds = null,
        IReadOnlyDictionary<int, string>? imageRelationshipOverrides = null,
        IReadOnlyDictionary<string, string>? oleRelationshipOverrides = null)
    {
        _document = document;
        _relationshipIds = relationshipIds;
        _imageRelationshipOverrides = imageRelationshipOverrides;
        _oleRelationshipOverrides = oleRelationshipOverrides;
        return this;
    }
    

    /// <summary>
    /// Builds a mapping from paragraph index to charts that should be emitted
    /// near that paragraph, based on ChartModel.ParagraphIndexHint. Charts
    /// whose hints are out of range are ignored here and will be handled by
    /// later fallback logic when needed.
    /// </summary>
    private static Dictionary<int, List<ChartModel>> BuildChartsByParagraphMap(DocumentModel document)
    {
        var map = new Dictionary<int, List<ChartModel>>();
        if (document.Charts == null || document.Charts.Count == 0)
            return map;

        int maxParagraphIndex = document.Paragraphs.Count > 0
            ? document.Paragraphs.Max(p => p.Index)
            : -1;

        foreach (var chart in document.Charts)
        {
            if (chart.ParagraphIndexHint < 0)
                continue;
            if (chart.ParagraphIndexHint > maxParagraphIndex)
                continue;

            if (!map.TryGetValue(chart.ParagraphIndexHint, out var list))
            {
                list = new List<ChartModel>();
                map[chart.ParagraphIndexHint] = list;
            }

            list.Add(chart);
        }

        return map;
    }

    /// <summary>
    /// Writes an inline chart reference for the given ChartModel using a
    /// standard wp:inline + a:graphic + c:chart structure.
    /// </summary>
    private void WriteChartInline(ChartModel chart, int chartIndex)
    {
        if (_document == null || _relationshipIds == null)
            return;

        // If we have no chart relationship block reserved, bail out.
        if (_relationshipIds.FirstChartRId <= 0)
            return;

        int relNumericId = _relationshipIds.FirstChartRId + chartIndex;
        if (relNumericId <= 0)
            return;

        string relId = $"rId{relNumericId}";

        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        // Reasonable default size for charts (~6x4 inches).
        int widthEmu = 5715000;
        int heightEmu = 3810000;

        _writer.WriteStartElement("w", "p", wNs);

        // Center the chart paragraph by default.
        _writer.WriteStartElement("w", "pPr", wNs);
        _writer.WriteStartElement("w", "jc", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "center");
        _writer.WriteEndElement(); // w:jc
        _writer.WriteEndElement(); // w:pPr

        _writer.WriteStartElement("w", "r", wNs);
        _writer.WriteStartElement("w", "drawing", wNs);

        _writer.WriteStartElement("wp", "inline", wpNs);
        _writer.WriteAttributeString("distT", "0");
        _writer.WriteAttributeString("distB", "0");
        _writer.WriteAttributeString("distL", "0");
        _writer.WriteAttributeString("distR", "0");

        _writer.WriteStartElement("wp", "extent", wpNs);
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement(); // wp:extent

        _writer.WriteStartElement("wp", "effectExtent", wpNs);
        _writer.WriteAttributeString("l", "0");
        _writer.WriteAttributeString("t", "0");
        _writer.WriteAttributeString("r", "0");
        _writer.WriteAttributeString("b", "0");
        _writer.WriteEndElement(); // wp:effectExtent

        // docPr with a simple name derived from the chart index or title.
        _writer.WriteStartElement("wp", "docPr", wpNs);
        _writer.WriteAttributeString("id", (1000 + chartIndex).ToString());
        var baseName = !string.IsNullOrEmpty(chart.Title)
            ? SanitizeXmlString(chart.Title)
            : $"Chart {chartIndex + 1}";
        _writer.WriteAttributeString("name", baseName);
        _writer.WriteEndElement(); // wp:docPr

        // Non-visual graphic frame properties.
        _writer.WriteStartElement("wp", "cNvGraphicFramePr", wpNs);
        _writer.WriteStartElement("a", "graphicFrameLocks", aNs);
        _writer.WriteAttributeString("noChangeAspect", "1");
        _writer.WriteEndElement(); // a:graphicFrameLocks
        _writer.WriteEndElement(); // wp:cNvGraphicFramePr

        // a:graphic / a:graphicData / c:chart
        _writer.WriteStartElement("a", "graphic", aNs);
        _writer.WriteStartElement("a", "graphicData", aNs);
        _writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart");

        _writer.WriteStartElement("c", "chart", cNs);
        _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", relId);
        _writer.WriteEndElement(); // c:chart

        _writer.WriteEndElement(); // a:graphicData
        _writer.WriteEndElement(); // a:graphic

        _writer.WriteEndElement(); // wp:inline
        _writer.WriteEndElement(); // w:drawing
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p
    }

    /// <summary>
    /// Writes the document content
    /// </summary>
    public void WriteDocument(DocumentModel document)
    {
        // start fresh for each document so track change IDs don't carry over
        _trackChangeId = 1;
        _bookmarkIds.Clear();
        _bookmarkCounter = 0;

        _document = document;
        _relationshipIds = RelationshipsWriter.ComputeRelationshipIds(document);
        
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "document", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // ensure there is at least a default section so writer has something to work with
        if (document.Properties.Sections.Count == 0)
        {
            document.Properties.Sections.Add(new SectionInfo());
        }
        
        // Add XML namespace definitions
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        _writer.WriteAttributeString("xmlns", "wp", null, "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("xmlns", "pic", null, "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteAttributeString("xmlns", "wps", null, "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        _writer.WriteAttributeString("xmlns", "v", null, "urn:schemas-microsoft-com:vml");
        _writer.WriteAttributeString("xmlns", "o", null, "urn:schemas-microsoft-com:office:office");
        _writer.WriteAttributeString("xmlns", "c", null, "http://schemas.openxmlformats.org/drawingml/2006/chart");
        
        WriteBody(document);
        
        _writer.WriteEndElement(); // w:document
        _writer.WriteEndDocument();
    }
    
    private void WriteBody(DocumentModel document)
    {
        _writer.WriteStartElement("w", "body", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Precompute section boundaries: which paragraph index ends each section
        var sectionEndMap = BuildSectionEndMap(document);

        // Track which paragraphs are part of tables
        var tableParagraphIndices = new HashSet<int>();
        foreach (var table in document.Tables)
        {
            for (int i = table.StartParagraphIndex; i <= table.EndParagraphIndex; i++)
            {
                tableParagraphIndices.Add(i);
            }
        }
        
        // Precompute shapes to emit near specific paragraphs and avoid duplicate images
        var shapesByParagraph = BuildShapesByParagraphMap(document, out var usedImageIndices);

        // Precompute charts to emit near specific paragraphs where we have
        // hints; charts without hints will be emitted near the end.
        var chartsByParagraph = BuildChartsByParagraphMap(document);

        // Build comment range mapping (annotation CP → paragraph index)
        BuildCommentRangeMap(document);

        // Suppress leading pageBreakBefore so first visible content (e.g. 绿色等级评价报告) appears on page 1
        _suppressLeadingPageBreak = true;

        // Write content: paragraphs and tables
        int paraIndex = 0;
        while (paraIndex < document.Paragraphs.Count)
        {
            // Check if this paragraph starts a table
            var table = document.Tables.FirstOrDefault(t => t.StartParagraphIndex == paraIndex);
            if (table != null)
            {
                WriteTable(table);
                _suppressLeadingPageBreak = false; // table is visible content

                // If a section ends at the last paragraph index of this table, emit sectPr here
                var lastParaOfTable = table.EndParagraphIndex;
                if (sectionEndMap.TryGetValue(lastParaOfTable, out var sectionForTable))
                {
                    if (!IsFinalSection(document, sectionForTable))
                    {
                        WriteSectionBreakParagraph(sectionForTable);
                    }
                }

                paraIndex = table.EndParagraphIndex + 1;
            }
            else
            {
                var paragraph = document.Paragraphs[paraIndex];

                // If a section ends at this paragraph, pass it so sectPr is embedded inside w:pPr
                SectionInfo? sectionForParagraph = null;
                sectionEndMap.TryGetValue(paragraph.Index, out sectionForParagraph);
                if (IsFinalSection(document, sectionForParagraph))
                {
                    sectionForParagraph = null;
                }

                WriteParagraph(paragraph, _suppressLeadingPageBreak, sectionForParagraph);
                if (_suppressLeadingPageBreak && ParagraphHasVisibleContent(paragraph))
                    _suppressLeadingPageBreak = false;

                // Emit any charts associated with this paragraph
                if (chartsByParagraph.TryGetValue(paragraph.Index, out var chartsForParagraph))
                {
                    foreach (var chart in chartsForParagraph)
                    {
                        WriteChartInline(chart, chart.Index);
                    }
                }

                // Emit any shapes that are associated with this paragraph
                if (shapesByParagraph.TryGetValue(paragraph.Index, out var shapesForParagraph))
                {
                    foreach (var shape in shapesForParagraph)
                    {
                        WriteInlinePictureShape(shape, document);
                    }
                }

                paraIndex++;
            }
        }

        // Write textboxes after main body content
        WriteTextboxes(document);
        
        WriteSections(document);
        
        _writer.WriteEndElement(); // w:body
    }

    /// <summary>
    /// Builds _commentStartsByParagraph and _commentEndsByParagraph by mapping
    /// annotation CP ranges to the paragraph that contains those CPs.
    /// </summary>
    private void BuildCommentRangeMap(DocumentModel document)
    {
        _commentStartsByParagraph.Clear();
        _commentEndsByParagraph.Clear();

        if (document.Annotations == null || document.Annotations.Count == 0)
            return;

        // Build a sorted list of (firstRunCp, paragraph index) for quick lookup.
        // Each entry represents the first CP of a paragraph. Empty paragraphs are assigned
        // a CP just past the previous paragraph to ensure comments anchored there stay put.
        var paragraphCpRanges = new List<(int startCp, int endCp, int paragraphIndex)>();
        int lastCp = 0;
        foreach (var para in document.Paragraphs)
        {
            int startCp;
            int endCp;
            if (para.Runs.Count == 0)
            {
                startCp = lastCp;
                endCp = lastCp;
            }
            else
            {
                startCp = para.Runs[0].CharacterPosition;
                var lastRun = para.Runs[para.Runs.Count - 1];
                endCp = lastRun.CharacterPosition + Math.Max(1, lastRun.CharacterLength);
            }
            paragraphCpRanges.Add((startCp, endCp, para.Index));
            if (endCp > lastCp) lastCp = endCp;
        }

        if (paragraphCpRanges.Count == 0) return;
        paragraphCpRanges.Sort((a, b) => a.startCp.CompareTo(b.startCp));

        // Helper: find the paragraph index for a given CP via binary search.
        int FindParagraphForCp(int cp)
        {
            int lo = 0, hi = paragraphCpRanges.Count - 1;
            int best = -1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (paragraphCpRanges[mid].startCp <= cp)
                {
                    best = mid;
                    lo = mid + 1;
                }
                else
                {
                    hi = mid - 1;
                }
            }
            if (best >= 0) return paragraphCpRanges[best].paragraphIndex;
            // Fallback: use the first paragraph
            return paragraphCpRanges[0].paragraphIndex;
        }

        for (int i = 0; i < document.Annotations.Count; i++)
        {
            var ann = document.Annotations[i];
            int commentId = i; // matches the ID written by CommentsWriter

            int startPara = FindParagraphForCp(ann.StartCharacterPosition);
            int endPara = ann.EndCharacterPosition > ann.StartCharacterPosition
                ? FindParagraphForCp(ann.EndCharacterPosition)
                : startPara;

            if (!_commentStartsByParagraph.TryGetValue(startPara, out var startList))
            {
                startList = new List<int>();
                _commentStartsByParagraph[startPara] = startList;
            }
            startList.Add(commentId);

            if (!_commentEndsByParagraph.TryGetValue(endPara, out var endList))
            {
                endList = new List<int>();
                _commentEndsByParagraph[endPara] = endList;
            }
            endList.Add(commentId);
        }
    }

    /// <summary>
    /// Builds a mapping from paragraph index to shapes that should be emitted
    /// near that paragraph, while also avoiding duplicate image indices that
    /// are already used elsewhere in the document.
    /// </summary>
    private Dictionary<int, List<ShapeModel>> BuildShapesByParagraphMap(DocumentModel document, out HashSet<int> usedImageIndices)
    {
        usedImageIndices = CollectUsedImageIndices(document);
        var map = new Dictionary<int, List<ShapeModel>>();

        if (document.Shapes == null || document.Shapes.Count == 0)
            return map;

        foreach (var shape in document.Shapes)
        {
            if (shape.ParagraphIndexHint < 0)
                continue;

            // 对于图片形状，我们需要避免重复：如果同一 imageIndex 已经作为正文
            // 图像出现过，就跳过这个形状；非图片矢量形状不参与去重。
            if (shape.Type == ShapeType.Picture && shape.ImageIndex is not null)
            {
                var imageIndex = shape.ImageIndex.Value;
                if (!usedImageIndices.Add(imageIndex))
                {
                    continue;
                }
            }

            if (!map.TryGetValue(shape.ParagraphIndexHint, out var list))
            {
                list = new List<ShapeModel>();
                map[shape.ParagraphIndexHint] = list;
            }

            list.Add(shape);
        }

        return map;
    }

    /// <summary>
    /// Collects all image indices that are already used in paragraphs, tables
    /// and textboxes so that we can avoid emitting duplicate images for shapes.
    /// </summary>
    private HashSet<int> CollectUsedImageIndices(DocumentModel document)
    {
        var used = new HashSet<int>();

        // Paragraph-level runs
        foreach (var para in document.Paragraphs)
        {
            foreach (var run in para.Runs)
            {
                if (run.IsPicture && run.ImageIndex >= 0)
                {
                    used.Add(run.ImageIndex);
                }
            }
        }

        // Tables
        foreach (var table in document.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        foreach (var run in para.Runs)
                        {
                            if (run.IsPicture && run.ImageIndex >= 0)
                            {
                                used.Add(run.ImageIndex);
                            }
                        }
                    }
                }
            }
        }

        // Textboxes
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs != null)
            {
                foreach (var para in textbox.Paragraphs)
                {
                    foreach (var run in para.Runs)
                    {
                        if (run.IsPicture && run.ImageIndex >= 0)
                        {
                            used.Add(run.ImageIndex);
                        }
                    }
                }
            }
        }

        return used;
    }

    
    private void WriteSections(DocumentModel document)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "sectPr", wNs);
        
        SectionInfo? lastSection = null;
        if (document.Properties.Sections.Count > 0)
            lastSection = document.Properties.Sections[document.Properties.Sections.Count - 1];

        WriteSectionContent(lastSection);
        _writer.WriteEndElement(); // sectPr
    }

    private void WriteSectionBreakParagraph(SectionInfo section)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "p", wNs);
        _writer.WriteStartElement("w", "pPr", wNs);
        _writer.WriteStartElement("w", "sectPr", wNs);
        WriteSectionContent(section);
        _writer.WriteEndElement(); // w:sectPr
        _writer.WriteEndElement(); // w:pPr
        _writer.WriteEndElement(); // w:p
    }
    
    private void WriteSectionProperties(DocumentProperties props)
    {
        // Legacy entry point kept for compatibility; delegate to the unified
        // section content writer using document-level properties.
        WriteSectionContent(null);
    }

    /// <summary>
    /// Writes the content of a w:sectPr element for either a specific section
    /// (SectionInfo) or, when section is null, for the document-level defaults.
    /// </summary>
    private void WriteSectionContent(SectionInfo? section)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        if (_document == null)
            return;

        var props = _document.Properties ?? new DocumentProperties();

        // headerReference and footerReference must come first in sectPr.
        // A sectPr can contain multiple references (default/first/even), and
        // these must be emitted together to activate first-page and facing-page
        // semantics in Word.
        int sectionIndex = section?.SectionIndex ?? -1;
        bool suppressHeaders = section?.HeaderReference == HeaderFooterReferenceType.None;
        bool suppressFooters = section?.FooterReference == HeaderFooterReferenceType.None;
        var defaultHeader = suppressHeaders ? null : FindHeaderFooter(_document.HeadersFooters.Headers, sectionIndex, HeaderFooterType.HeaderOdd);
        var firstHeader = suppressHeaders ? null : FindHeaderFooter(_document.HeadersFooters.Headers, sectionIndex, HeaderFooterType.HeaderFirst);
        var evenHeader = suppressHeaders ? null : FindHeaderFooter(_document.HeadersFooters.Headers, sectionIndex, HeaderFooterType.HeaderEven);
        var defaultFooter = suppressFooters ? null : FindHeaderFooter(_document.HeadersFooters.Footers, sectionIndex, HeaderFooterType.FooterOdd);
        var firstFooter = suppressFooters ? null : FindHeaderFooter(_document.HeadersFooters.Footers, sectionIndex, HeaderFooterType.FooterFirst);
        var evenFooter = suppressFooters ? null : FindHeaderFooter(_document.HeadersFooters.Footers, sectionIndex, HeaderFooterType.FooterEven);
        bool allowsHeaders = defaultHeader != null || firstHeader != null || evenHeader != null;
        bool allowsFooters = defaultFooter != null || firstFooter != null || evenFooter != null;
        bool usesFirstPage = section?.TitlePage == true || firstHeader != null || firstFooter != null;
        bool usesEvenAndOdd = UsesEvenAndOddHeaders(_document) && (evenHeader != null || evenFooter != null);

        if (allowsHeaders)
        {
            WriteHeaderFooterReference("headerReference", "default", defaultHeader?.RelationshipId);
            if (usesFirstPage)
                WriteHeaderFooterReference("headerReference", "first", firstHeader?.RelationshipId);
            if (usesEvenAndOdd)
                WriteHeaderFooterReference("headerReference", "even", evenHeader?.RelationshipId);
        }

        if (allowsFooters)
        {
            WriteHeaderFooterReference("footerReference", "default", defaultFooter?.RelationshipId);
            if (usesFirstPage)
                WriteHeaderFooterReference("footerReference", "first", firstFooter?.RelationshipId);
            if (usesEvenAndOdd)
                WriteHeaderFooterReference("footerReference", "even", evenFooter?.RelationshipId);
        }

        if (section?.BreakCode is >= 0 and <= 4)
        {
            var sectionType = section.BreakCode switch
            {
                0 => "continuous",
                1 => "nextColumn",
                2 => "nextPage",
                3 => "evenPage",
                4 => "oddPage",
                _ => null
            };
            if (!string.IsNullOrEmpty(sectionType))
            {
                _writer.WriteStartElement("w", "type", wNs);
                _writer.WriteAttributeString("w", "val", wNs, sectionType);
                _writer.WriteEndElement();
            }
        }

        if (usesFirstPage)
        {
            _writer.WriteStartElement("w", "titlePg", wNs);
            _writer.WriteEndElement();
        }

        // Page size and margins: prefer per-section overrides when available
        // pgSz
        _writer.WriteStartElement("w", "pgSz", wNs);
        int w = ClampTwips(section?.PageWidth > 0 ? section.PageWidth : props.PageWidth, 720, 31680, 12240);
        int h = ClampTwips(section?.PageHeight > 0 ? section.PageHeight : props.PageHeight, 720, 31680, 15840);
        _writer.WriteAttributeString("w", "w", wNs, w.ToString());
        _writer.WriteAttributeString("w", "h", wNs, h.ToString());
        if (section?.IsLandscape == true || (section == null && props.IsLandscape))
            _writer.WriteAttributeString("w", "orient", wNs, "landscape");
        _writer.WriteEndElement();

        // pgMar
        _writer.WriteStartElement("w", "pgMar", wNs);
        _writer.WriteAttributeString("w", "top", wNs, ClampTwips(section?.MarginTop != 0 ? section!.MarginTop : props.MarginTop, 0, 15840, 1440).ToString());
        _writer.WriteAttributeString("w", "right", wNs, ClampTwips(section?.MarginRight != 0 ? section!.MarginRight : props.MarginRight, 0, 15840, 1440).ToString());
        _writer.WriteAttributeString("w", "bottom", wNs, ClampTwips(section?.MarginBottom != 0 ? section!.MarginBottom : props.MarginBottom, 0, 15840, 1440).ToString());
        _writer.WriteAttributeString("w", "left", wNs, ClampTwips(section?.MarginLeft != 0 ? section!.MarginLeft : props.MarginLeft, 0, 15840, 1440).ToString());
        _writer.WriteAttributeString("w", "header", wNs, ClampTwips(section?.HeaderMargin ?? (section == null ? 720 : 0), 0, 15840, 720).ToString());
        _writer.WriteAttributeString("w", "footer", wNs, ClampTwips(section?.FooterMargin ?? (section == null ? 720 : 0), 0, 15840, 720).ToString());
        _writer.WriteAttributeString("w", "gutter", wNs, ClampTwips(section?.Gutter ?? 0, 0, 31680, 0).ToString());
        _writer.WriteEndElement();

        // Mirror margins (left/right swapped on facing pages) – driven by DOP flag.
        if (props.FMirrorMargins)
        {
            _writer.WriteStartElement("w", "mirrorMargins", wNs);
            // 'val' attribute is optional per spec; omit to avoid redundant data.
            _writer.WriteEndElement();
        }

        // Page numbering start (document-level only, for now)
        if (section?.PageNumberStart is int pageNumberStart)
        {
            _writer.WriteStartElement("w", "pgNumType", wNs);
            _writer.WriteAttributeString("w", "start", wNs, pageNumberStart.ToString());
            _writer.WriteEndElement();
        }
        else if (section == null && props.SectionStartPageNumber > 1)
        {
            _writer.WriteStartElement("w", "pgNumType", wNs);
            _writer.WriteAttributeString("w", "start", wNs, props.SectionStartPageNumber.ToString());
            _writer.WriteEndElement();
        }

        // Columns
        _writer.WriteStartElement("w", "cols", wNs);
        int columnCount = section?.ColumnCount > 0 ? section.ColumnCount : 1;
        int columnSpacing = section?.ColumnSpacing > 0 ? section.ColumnSpacing : 720;
        if (columnCount > 1)
            _writer.WriteAttributeString("w", "num", wNs, columnCount.ToString());
        _writer.WriteAttributeString("w", "space", wNs, columnSpacing.ToString());
        _writer.WriteEndElement();

        if (section?.VerticalAlignment > 0)
        {
            var verticalAlignment = section.VerticalAlignment switch
            {
                1 => "center",
                2 => "both",
                3 => "bottom",
                _ => null
            };
            if (!string.IsNullOrEmpty(verticalAlignment))
            {
                _writer.WriteStartElement("w", "vAlign", wNs);
                _writer.WriteAttributeString("w", "val", wNs, verticalAlignment);
                _writer.WriteEndElement();
            }
        }

        if (section?.DocGridLinePitch > 0)
        {
            _writer.WriteStartElement("w", "docGrid", wNs);
            _writer.WriteAttributeString("w", "type", wNs, "lines");
            _writer.WriteAttributeString("w", "linePitch", wNs, section.DocGridLinePitch.ToString());
            _writer.WriteEndElement();
        }
    }

    private void WriteHeaderFooterReference(string elementName, string type, string? relationshipId)
    {
        if (string.IsNullOrEmpty(relationshipId))
            return;

        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", elementName, wNs);
        _writer.WriteAttributeString("w", "type", wNs, type);
        _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", relationshipId);
        _writer.WriteEndElement();
    }

    private static bool UsesEvenAndOddHeaders(DocumentModel document)
    {
        return document.Properties.FFacingPages ||
               document.HeadersFooters.Headers.Any(h => h.Type == HeaderFooterType.HeaderEven && HeaderFooterContentHelper.HasUsableContent(h)) ||
               document.HeadersFooters.Footers.Any(f => f.Type == HeaderFooterType.FooterEven && HeaderFooterContentHelper.HasUsableContent(f));
    }

    /// <summary>
    /// Builds a map from the paragraph index that ends each section to the
    /// corresponding SectionInfo, based on Sections[i].StartParagraphIndex.
    /// </summary>
    private static Dictionary<int, SectionInfo> BuildSectionEndMap(DocumentModel document)
    {
        var map = new Dictionary<int, SectionInfo>();
        var sections = document.Properties.Sections;
        if (sections.Count == 0 || document.Paragraphs.Count == 0)
            return map;

        for (int i = 0; i < sections.Count; i++)
        {
            var section = sections[i];
            int start = Math.Clamp(section.StartParagraphIndex, 0, document.Paragraphs.Count - 1);
            int end;

            if (i + 1 < sections.Count)
            {
                // This section ends just before the next section's start
                var nextStart = Math.Clamp(sections[i + 1].StartParagraphIndex, 0, document.Paragraphs.Count - 1);
                end = Math.Clamp(nextStart - 1, start, document.Paragraphs.Count - 1);
            }
            else
            {
                // Last section ends at the last paragraph
                end = document.Paragraphs.Count - 1;
            }

            if (!map.ContainsKey(end))
            {
                map[end] = section;
            }
        }

        return map;
    }

    private static bool IsFinalSection(DocumentModel document, SectionInfo? section)
    {
        if (section == null || document.Properties.Sections.Count == 0)
            return false;

        if (section.SectionIndex >= 0)
            return section.SectionIndex == document.Properties.Sections.Count - 1;

        return ReferenceEquals(section, document.Properties.Sections[document.Properties.Sections.Count - 1]);
    }

    private static HeaderFooterModel? FindHeaderFooter(IReadOnlyList<HeaderFooterModel> items, int sectionIndex, HeaderFooterType type)
    {
        HeaderFooterModel? exactMatch = null;
        HeaderFooterModel? sharedMatch = null;
        HeaderFooterModel? fallbackMatch = null;

        foreach (var item in items)
        {
            if (item.Type != type || !HeaderFooterContentHelper.HasUsableContent(item))
                continue;

            if (item.SectionIndex == sectionIndex)
                return item;

            if (item.SectionIndex < 0 && sharedMatch == null)
                sharedMatch = item;

            fallbackMatch ??= item;
            exactMatch ??= item;
        }

        return sharedMatch ?? exactMatch ?? fallbackMatch;
    }

    
    private void WriteSectionPropertiesCore(SectionInfo section)
    {
        _writer.WriteStartElement("w", "pgSz", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", section.PageWidth.ToString());
        _writer.WriteAttributeString("w", "h", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", section.PageHeight.ToString());
        if (section.IsLandscape)
        {
            _writer.WriteAttributeString("w", "orient", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "landscape");
        }
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "pgMar", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "top", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", section.MarginTop.ToString());
        _writer.WriteAttributeString("w", "right", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", section.MarginRight.ToString());
        _writer.WriteAttributeString("w", "bottom", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", section.MarginBottom.ToString());
        _writer.WriteAttributeString("w", "left", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", section.MarginLeft.ToString());
        _writer.WriteEndElement();
    }
    
    private void WriteSectionProperties()
    {
        _writer.WriteStartElement("w", "cols", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "space", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "720");
        _writer.WriteEndElement();
    }
    
    private static bool ParagraphHasVisibleContent(ParagraphModel paragraph)
    {
        return paragraph.Runs != null && paragraph.Runs.Any(r =>
            (!string.IsNullOrEmpty(r.Text) && !string.IsNullOrWhiteSpace(r.Text)) || r.IsPicture || r.IsField || r.IsBookmark);
    }

    /// <summary>
    /// Clamps twips-like metrics to a safe range to prevent invalid OOXML values
    /// from triggering Word repair mode on edge/corrupt source inputs.
    /// </summary>
    private static int ClampTwips(int value, int min, int max, int fallback)
    {
        if (value == 0)
            return fallback;
        return Math.Clamp(value, min, max);
    }

    private static int ConvertCharacterIndentToTwips(int charIndent, int fontSizeHalfPoints)
    {
        if (charIndent == 0)
            return 0;

        int effectiveFontSizeHalfPoints = fontSizeHalfPoints > 0 ? fontSizeHalfPoints : 24;
        return (int)Math.Round(Math.Abs(charIndent) * effectiveFontSizeHalfPoints * 10d / 100d, MidpointRounding.AwayFromZero);
    }

    private int ResolveParagraphIndentFontSizeHalfPoints(ParagraphModel paragraph)
    {
        foreach (var run in paragraph.Runs)
        {
            if (run.Properties == null)
                continue;

            int runFontSize = Math.Max(run.Properties.FontSize, run.Properties.FontSizeCs);
            if (runFontSize > 0)
                return runFontSize;
        }

        if (_document?.Styles?.Styles != null && paragraph.Properties?.StyleIndex > 0)
        {
            var style = _document.Styles.Styles.FirstOrDefault(s =>
                s.Type == StyleType.Paragraph && s.StyleId == paragraph.Properties.StyleIndex);
            if (style?.RunProperties != null)
            {
                int styleFontSize = Math.Max(style.RunProperties.FontSize, style.RunProperties.FontSizeCs);
                if (styleFontSize > 0)
                    return styleFontSize;
            }
        }

        return 24;
    }

    /// <summary>
    /// Writes a single paragraph, optionally suppressing a leading page break or
    /// appending a section break to the paragraph properties.
    /// </summary>
    /// <param name="paragraph">The paragraph model to write.</param>
    /// <param name="suppressPageBreakBefore">Whether a pageBreakBefore flag should be ignored for this paragraph.</param>
    /// <param name="sectionBreak">Optional section properties to append at the end of the paragraph properties.</param>
    public void WriteParagraph(ParagraphModel paragraph, bool suppressPageBreakBefore = false, SectionInfo? sectionBreak = null)
    {
        // If this paragraph is actually a wrapper for a nested table, write the table directly
        if (paragraph.Type == ParagraphType.NestedTable && paragraph.NestedTable != null)
        {
            WriteTable(paragraph.NestedTable);
            return;
        }

        // Filter runs to only those with actual content
        var runsWithContent = paragraph.Runs.Where(HasRenderableContent).ToList();
        
        // Always write the paragraph element - OOXML requires at least one w:p in table cells,
        // and empty paragraphs (blank lines, page breaks) are meaningful document structure.
        _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        WriteParagraphProperties(paragraph, suppressPageBreakBefore, sectionBreak);

        // Emit w:commentRangeStart for any comments that start at this paragraph
        if (_commentStartsByParagraph.TryGetValue(paragraph.Index, out var commentStarts))
        {
            foreach (var commentId in commentStarts)
            {
                _writer.WriteStartElement("w", "commentRangeStart", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", commentId.ToString());
                _writer.WriteEndElement();
            }
        }
        
        foreach (var run in runsWithContent)
        {
            WriteRun(run);
        }

        // Emit w:commentRangeEnd and w:r > w:commentReference for any comments that end at this paragraph
        if (_commentEndsByParagraph.TryGetValue(paragraph.Index, out var commentEnds))
        {
            foreach (var commentId in commentEnds)
            {
                _writer.WriteStartElement("w", "commentRangeEnd", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", commentId.ToString());
                _writer.WriteEndElement();

                // w:r containing w:commentReference
                _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "rPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "rStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "CommentReference");
                _writer.WriteEndElement();
                _writer.WriteEndElement(); // w:rPr
                _writer.WriteStartElement("w", "commentReference", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", commentId.ToString());
                _writer.WriteEndElement();
                _writer.WriteEndElement(); // w:r
            }
        }
        
        _writer.WriteEndElement(); // w:p
    }
    
    private void WriteParagraphProperties(ParagraphModel paragraph, bool suppressPageBreakBefore = false, SectionInfo? sectionBreak = null)
    {
        var props = paragraph.Properties;

        // Always emit w:pPr if there is a sectionBreak, even when props is null
        if (props == null && sectionBreak == null) return;
        
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        
        // pPr sequence per ISO 29500 CT_PPr:
        // pStyle -> keepNext -> keepLines -> pageBreakBefore -> numPr -> pBdr -> shd ->
        // spacing -> ind -> jc -> outlineLvl
        _writer.WriteStartElement("w", "pPr", wNs);
        
        // 1. pStyle
        if (props != null && props.StyleIndex >= 0)
        {
            var style = _document?.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == props.StyleIndex);
            var styleId = StyleHelper.GetParagraphStyleId(props.StyleIndex, style?.Name);
            
            _writer.WriteStartElement("w", "pStyle", wNs);
            _writer.WriteAttributeString("w", "val", wNs, styleId);
            _writer.WriteEndElement();
        }

        // 2. keepNext
        if (props != null && props.KeepWithNext)
        {
            _writer.WriteStartElement("w", "keepNext", wNs);
            _writer.WriteEndElement();
        }
        
        // 3. keepLines
        if (props != null && props.KeepTogether)
        {
            _writer.WriteStartElement("w", "keepLines", wNs);
            _writer.WriteEndElement();
        }
        
        // 4. pageBreakBefore (suppressed at doc start so first content e.g. 绿色等级评价报告 stays on page 1)
        if (props != null && props.PageBreakBefore && !suppressPageBreakBefore)
        {
            _writer.WriteStartElement("w", "pageBreakBefore", wNs);
            _writer.WriteEndElement();
        }

        // 5. numPr
        if (props != null && props.ListFormatId > 0)
        {
            WriteNumberingProperties(props.ListFormatId, props.ListLevel);
        }

        // 6. pBdr
        if (props != null && (props.BorderTop != null || props.BorderBottom != null || 
            props.BorderLeft != null || props.BorderRight != null))
        {
            _writer.WriteStartElement("w", "pBdr", wNs);
            if (props.BorderTop != null) WriteBorder("top", props.BorderTop);
            if (props.BorderBottom != null) WriteBorder("bottom", props.BorderBottom);
            if (props.BorderLeft != null) WriteBorder("left", props.BorderLeft);
            if (props.BorderRight != null) WriteBorder("right", props.BorderRight);
            _writer.WriteEndElement();
        }
        
        // 7. shd
        if (props != null && props.Shading != null)
        {
            WriteShading(props.Shading);
        }

        // 8. spacing
        bool hasExplicitLineSpacing = props != null &&
            (props.HasExplicitLineSpacing || props.LineSpacing != 240 || props.LineSpacingMultiple != 1);
        if (props != null && (props.SpaceBefore > 0 || props.SpaceBeforeLines > 0 || props.SpaceAfter > 0 || props.SpaceAfterLines > 0 || hasExplicitLineSpacing))
        {
            _writer.WriteStartElement("w", "spacing", wNs);
            if (props.SpaceBeforeLines > 0)
                _writer.WriteAttributeString("w", "beforeLines", wNs, props.SpaceBeforeLines.ToString());
            else if (props.SpaceBefore > 0)
                _writer.WriteAttributeString("w", "before", wNs, props.SpaceBefore.ToString());
            if (props.SpaceAfterLines > 0)
                _writer.WriteAttributeString("w", "afterLines", wNs, props.SpaceAfterLines.ToString());
            else if (props.SpaceAfter > 0)
                _writer.WriteAttributeString("w", "after", wNs, props.SpaceAfter.ToString());
            if (hasExplicitLineSpacing)
            {
                // In MS-DOC LSPD: fMultLinespace=1 means proportional (auto),
                // fMultLinespace=0 means absolute. Negative dyaLine = exact,
                // positive dyaLine with fMult=0 = atLeast.
                int lineVal = props.LineSpacing;
                string lineRule;
                if (props.LineSpacingMultiple == 1)
                {
                    // Proportional: value is in 240ths of a line (240 = single)
                    lineRule = "auto";
                }
                else if (lineVal < 0)
                {
                    // Exact spacing: use absolute value
                    lineVal = Math.Abs(lineVal);
                    lineRule = "exact";
                }
                else
                {
                    // At-least spacing
                    lineRule = "atLeast";
                }
                _writer.WriteAttributeString("w", "line", wNs, lineVal.ToString());
                _writer.WriteAttributeString("w", "lineRule", wNs, lineRule);
            }
            _writer.WriteEndElement();
        }
        
        // 9. ind
        if (props != null && (props.IndentLeft != 0 || props.IndentLeftChars != 0 || props.IndentRight != 0 || props.IndentRightChars != 0 || props.IndentFirstLine != 0 || props.IndentFirstLineChars != 0))
        {
            int fontSizeHalfPoints = ResolveParagraphIndentFontSizeHalfPoints(paragraph);
            _writer.WriteStartElement("w", "ind", wNs);
            int indentLeft = props.IndentLeft != 0
                ? props.IndentLeft
                : ConvertCharacterIndentToTwips(props.IndentLeftChars, fontSizeHalfPoints);
            if (indentLeft != 0)
                _writer.WriteAttributeString("w", "left", wNs, indentLeft.ToString());
            if (props.IndentLeftChars != 0)
                _writer.WriteAttributeString("w", "leftChars", wNs, props.IndentLeftChars.ToString());
            int indentRight = props.IndentRight != 0
                ? props.IndentRight
                : ConvertCharacterIndentToTwips(props.IndentRightChars, fontSizeHalfPoints);
            if (indentRight != 0)
                _writer.WriteAttributeString("w", "right", wNs, indentRight.ToString());
            if (props.IndentRightChars != 0)
                _writer.WriteAttributeString("w", "rightChars", wNs, props.IndentRightChars.ToString());
            
            if (props.IndentFirstLineChars > 0)
            {
                int firstLine = props.IndentFirstLine > 0
                    ? props.IndentFirstLine
                    : ConvertCharacterIndentToTwips(props.IndentFirstLineChars, fontSizeHalfPoints);
                if (firstLine > 0)
                    _writer.WriteAttributeString("w", "firstLine", wNs, firstLine.ToString());
                _writer.WriteAttributeString("w", "firstLineChars", wNs, props.IndentFirstLineChars.ToString());
            }
            else if (props.IndentFirstLineChars < 0)
            {
                int hanging = props.IndentFirstLine < 0
                    ? Math.Abs(props.IndentFirstLine)
                    : ConvertCharacterIndentToTwips(props.IndentFirstLineChars, fontSizeHalfPoints);
                if (hanging > 0)
                    _writer.WriteAttributeString("w", "hanging", wNs, hanging.ToString());
                _writer.WriteAttributeString("w", "hangingChars", wNs, Math.Abs(props.IndentFirstLineChars).ToString());
            }
            else if (props.IndentFirstLine > 0)
            {
                _writer.WriteAttributeString("w", "firstLine", wNs, props.IndentFirstLine.ToString());
            }
            else if (props.IndentFirstLine < 0)
            {
                _writer.WriteAttributeString("w", "hanging", wNs, Math.Abs(props.IndentFirstLine).ToString());
            }
            _writer.WriteEndElement();
        }
        
        // 10. jc
        if (props != null && props.Alignment != ParagraphAlignment.Left)
        {
            _writer.WriteStartElement("w", "jc", wNs);
            var alignment = props.Alignment switch
            {
                ParagraphAlignment.Center => "center",
                ParagraphAlignment.Right => "right",
                ParagraphAlignment.Justify => "both",
                ParagraphAlignment.Distributed => "distribute",
                _ => "left"
            };
            _writer.WriteAttributeString("w", "val", wNs, alignment);
            _writer.WriteEndElement();
        }

        // 11. outlineLvl
        if (props != null && props.OutlineLevel >= 0 && props.OutlineLevel < 9)
        {
            _writer.WriteStartElement("w", "outlineLvl", wNs);
            _writer.WriteAttributeString("w", "val", wNs, props.OutlineLevel.ToString());
            _writer.WriteEndElement();
        }

        // 12. Text Formatting / Typography Flags
        if (props != null && !props.WordWrap)
        {
            _writer.WriteStartElement("w", "wordWrap", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (props != null && !props.Kinsoku)
        {
            _writer.WriteStartElement("w", "kinsoku", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (props != null && !props.SnapToGrid)
        {
            _writer.WriteStartElement("w", "snapToGrid", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (props != null && !props.AutoSpaceDe)
        {
            _writer.WriteStartElement("w", "autoSpaceDE", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (props != null && !props.AutoSpaceDn)
        {
            _writer.WriteStartElement("w", "autoSpaceDN", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (props != null && props.TopLinePunct)
        {
            _writer.WriteStartElement("w", "topLinePunct", wNs);
            _writer.WriteEndElement();
        }
        if (props != null && props.OverflowPunct)
        {
            _writer.WriteStartElement("w", "overflowPunct", wNs);
            _writer.WriteEndElement();
        }
        // Non-final section break: must be the last child of w:pPr per OOXML spec
        if (sectionBreak != null)
        {
            _writer.WriteStartElement("w", "sectPr", wNs);
            WriteSectionContent(sectionBreak);
            _writer.WriteEndElement();
        }
        
        _writer.WriteEndElement(); // w:pPr
    }

    /// <summary>
    /// Writes numbering properties (w:numPr) for list paragraphs.
    /// OOXML CT_NumPr order: ilvl, numId
    /// </summary>
    private void WriteNumberingProperties(int listFormatId, int listLevel)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "numPr", wNs);
        
        // 1. ilvl (must come before numId per schema)
        _writer.WriteStartElement("w", "ilvl", wNs);
        _writer.WriteAttributeString("w", "val", wNs, listLevel.ToString());
        _writer.WriteEndElement();
        
        // 2. numId
        _writer.WriteStartElement("w", "numId", wNs);
        _writer.WriteAttributeString("w", "val", wNs, listFormatId.ToString());
        _writer.WriteEndElement();
        
        _writer.WriteEndElement(); // w:numPr
    }
    
    private void WriteBorder(string position, BorderInfo border)
    {
        if (border.Style == BorderStyle.None || IsLikelyMalformedBorder(border)) return;
        
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        string? themeColor = ColorHelper.GetThemeColorName(border.Color);
        string? resolvedThemeHex = ColorHelper.ResolveThemeColorHex(border.Color, _document?.Theme);
        _writer.WriteStartElement("w", position, wNs);
        _writer.WriteAttributeString("w", "val", wNs, GetBorderStyle(border.Style));
        // Width is in 1/8 pt (same as OOXML w:sz units) after BRC80 decode
        _writer.WriteAttributeString("w", "sz", wNs, border.Width.ToString());
        _writer.WriteAttributeString("w", "space", wNs, border.Space.ToString());
        _writer.WriteAttributeString("w", "color", wNs, resolvedThemeHex ?? ColorHelper.ResolveColorHex(border.Color, _document?.Theme));
        if (themeColor != null)
        {
            _writer.WriteAttributeString("w", "themeColor", wNs, themeColor);
        }
        _writer.WriteEndElement();
    }

    private static bool IsLikelyMalformedBorder(BorderInfo border)
    {
        return border.Width > 96 && border.Color == 255;
    }
    
    private void WriteShading(ShadingInfo shading)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        string? foregroundThemeColor = ColorHelper.GetThemeColorName(shading.ForegroundColor);
        string? foregroundThemeHex = ColorHelper.ResolveThemeColorHex(shading.ForegroundColor, _document?.Theme);
        string? backgroundThemeColor = ColorHelper.GetThemeColorName(shading.BackgroundColor);
        string? backgroundThemeHex = ColorHelper.ResolveThemeColorHex(shading.BackgroundColor, _document?.Theme);
        _writer.WriteStartElement("w", "shd", wNs);
        // Use PatternVal (from SHD ipat) when set; otherwise map Pattern enum to OOXML val so pattern/tiled background is preserved
        var val = !string.IsNullOrEmpty(shading.PatternVal)
            ? shading.PatternVal
            : ShadingPatternToShdVal(shading.Pattern);
        _writer.WriteAttributeString("w", "val", wNs, val);
        if (shading.ForegroundColor != 0)
        {
            _writer.WriteAttributeString("w", "color", wNs, foregroundThemeHex ?? ColorHelper.ResolveColorHex(shading.ForegroundColor, _document?.Theme));
            if (foregroundThemeColor != null)
            {
                _writer.WriteAttributeString("w", "themeColor", wNs, foregroundThemeColor);
            }
        }
        _writer.WriteAttributeString("w", "fill", wNs, backgroundThemeHex ?? ColorHelper.ResolveColorHex(shading.BackgroundColor, _document?.Theme, fallback: "FFFFFF"));
        if (backgroundThemeColor != null)
        {
            _writer.WriteAttributeString("w", "themeFill", wNs, backgroundThemeColor);
        }
        _writer.WriteEndElement();
    }

    private static string ShadingPatternToShdVal(ShadingPattern pattern)
    {
        return pattern switch
        {
            ShadingPattern.Clear => "clear",
            ShadingPattern.Solid => "solid",
            ShadingPattern.Percent5 => "pct5",
            ShadingPattern.Percent10 => "pct10",
            ShadingPattern.Percent20 => "pct20",
            ShadingPattern.Percent25 => "pct25",
            ShadingPattern.Percent30 => "pct30",
            ShadingPattern.Percent40 => "pct40",
            ShadingPattern.Percent50 => "pct50",
            ShadingPattern.Percent60 => "pct60",
            ShadingPattern.Percent70 => "pct70",
            ShadingPattern.Percent75 => "pct75",
            ShadingPattern.Percent80 => "pct80",
            ShadingPattern.Percent90 => "pct90",
            ShadingPattern.LightHorizontal => "thinHorzStripe",
            ShadingPattern.DarkHorizontal => "horzStripe",
            ShadingPattern.LightVertical => "thinVertStripe",
            ShadingPattern.DarkVertical => "vertStripe",
            ShadingPattern.LightDiagonalDown => "thinDiagStripe",
            ShadingPattern.LightDiagonalUp => "thinReverseDiagStripe",
            ShadingPattern.DarkDiagonalDown => "diagStripe",
            ShadingPattern.DarkDiagonalUp => "reverseDiagStripe",
            ShadingPattern.DarkGrid => "horzCross",
            ShadingPattern.DarkTrellis => "diagCross",
            ShadingPattern.LightGray => "pct25",
            ShadingPattern.MediumGray => "pct50",
            ShadingPattern.DarkGray => "pct75",
            _ => "clear"
        };
    }
    
    
    private string GetBorderStyle(BorderStyle style)
    {
        return style switch
        {
            BorderStyle.Single => "single",
            BorderStyle.Thick => "thick",
            BorderStyle.Double => "double",
            BorderStyle.Dotted => "dotted",
            BorderStyle.Dashed => "dash",
            BorderStyle.DotDash => "dotDash",
            BorderStyle.DotDotDash => "dotDotDash",
            BorderStyle.Wave => "wave",
            _ => "nil"
        };
    }
    
    private void WriteTrackChangeStart(string type, RunProperties props)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", type, wNs);
        _writer.WriteAttributeString("w", "id", wNs, (_trackChangeId++).ToString());
        
        string author = "Unknown Author";
        ushort authorIdx = type == "ins" ? props.AuthorIndexIns : props.AuthorIndexDel;
        if (_document != null && authorIdx < _document.RevisionAuthors.Count)
        {
            author = _document.RevisionAuthors[authorIdx];
        }
        _writer.WriteAttributeString("w", "author", wNs, author);
        
        uint dttm = type == "ins" ? props.DateIns : props.DateDel;
        if (dttm != 0)
        {
            var dt = DttmHelper.ParseDttm(dttm);
            _writer.WriteAttributeString("w", "date", wNs, dt.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }
    }

    private void WriteRun(RunModel run)
    {
        // Skip runs with no content at all (no text, no picture, no field)
        bool hasText = !string.IsNullOrEmpty(run.Text);
        bool hasVisualContent = run.IsPicture || run.IsField;
        
        if (!hasText && !hasVisualContent)
        {
            if (run.IsBookmark)
            {
                if (run.IsBookmarkStart)
                {
                    WriteBookmarkStart(run.BookmarkName);
                }
                else
                {
                    WriteBookmarkEnd(run.BookmarkName);
                }
            }

            // Even if no text, if there are properties, we might want to write them
            // But for now, skip empty runs to avoid corruption
            return;
        }

        bool isInserted = run.Properties != null && run.Properties.IsInserted;
        bool isDeleted = run.Properties != null && run.Properties.IsDeleted;

        if (isInserted) WriteTrackChangeStart("ins", run.Properties!);
        if (isDeleted) WriteTrackChangeStart("del", run.Properties!);

        // Handle bookmark start
        if (run.IsBookmark && run.IsBookmarkStart)
        {
            WriteBookmarkStart(run.BookmarkName);
        }

        // Handle hyperlink (skip entirely when hyperlinks are disabled)
        if (_options.EnableHyperlinks && run.IsHyperlink && RunHasHyperlinkTarget(run))
        {
            // if the run text contains extra material before an embedded HYPERLINK
            // field code (common when the reader leaves the field code in the same
            // run as preceding Chinese text), split the run so that the prefix is
            // written as an ordinary run and the remaining portion is treated as
            // the hyperlink.  this prevents stray non-link text from appearing
            // inside the w:hyperlink element and allows sanitization to drop the
            // field code more reliably.
            if (!string.IsNullOrEmpty(run.Text))
            {
                int idx = run.Text.IndexOf("HYPERLINK", StringComparison.OrdinalIgnoreCase);
                if (idx > 0)
                {
                    string before = run.Text.Substring(0, idx);
                    string after = run.Text.Substring(idx);
                    // write the prefix as a normal run with the same formatting
                    var prefix = new RunModel
                    {
                        Text = before,
                        Properties = run.Properties,
                        IsField = run.IsField,
                        FieldCode = run.FieldCode,
                        CharacterPosition = run.CharacterPosition,
                        CharacterLength = run.CharacterLength,
                        IsPicture = run.IsPicture,
                        ImageIndex = run.ImageIndex,
                        DisplayWidthTwips = run.DisplayWidthTwips,
                        DisplayHeightTwips = run.DisplayHeightTwips,
                        FcPic = run.FcPic,
                        ImageRelationshipId = run.ImageRelationshipId,
                        IsOle = run.IsOle,
                        OleObjectId = run.OleObjectId,
                        OleProgId = run.OleProgId,
                        // explicitly clear hyperlink flags
                        IsHyperlink = false,
                        HyperlinkUrl = null,
                        HyperlinkBookmark = null,
                        HyperlinkRelationshipId = null,
                        IsBookmark = run.IsBookmark,
                        BookmarkName = run.BookmarkName,
                        IsBookmarkStart = run.IsBookmarkStart,
                        CropTop = run.CropTop,
                        CropBottom = run.CropBottom,
                        CropLeft = run.CropLeft,
                        CropRight = run.CropRight,
                        FlipHorizontal = run.FlipHorizontal,
                        FlipVertical = run.FlipVertical
                    };
                    WriteRun(prefix);
                    run.Text = after; // continue processing remaining text below
                }
            }

            WriteHyperlink(run);
        }
        else
        {
            _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            WriteRunProperties(run);

            if ((run.IsPicture && run.ImageIndex >= 0) || (run.IsOle && !string.IsNullOrEmpty(run.OleObjectId) && !string.IsNullOrEmpty(run.OleProgId)))
            {
                if (run.IsOle && !string.IsNullOrEmpty(run.OleObjectId) && !string.IsNullOrEmpty(run.OleProgId))
                {
                    var oleObj = _document?.OleObjects.FirstOrDefault(o => o.ObjectId == run.OleObjectId);
                    if (oleObj != null && !string.IsNullOrEmpty(oleObj.MathContent))
                    {
                        // Native Math (OMML) should be a sibling of w:r, not inside it.
                        _writer.WriteEndElement(); // w:r
                        _writer.WriteRaw(oleObj.MathContent);
                    }
                    else
                    {
                        WriteOleObject(run);
                        _writer.WriteEndElement(); // w:r
                    }
                }
                else
                {
                    WritePicture(run);
                    _writer.WriteEndElement(); // w:r
                }
            }
            else if (run.IsField)
            {
                // OOXML requires fldChar begin/separate/end in separate w:r elements
                // Run 1: fldChar begin
                _writer.WriteStartElement("w", "fldChar", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "fldCharType", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "begin");
                _writer.WriteEndElement();
                _writer.WriteEndElement(); // w:r (begin)

                // Run 2: instrText
                if (!string.IsNullOrEmpty(run.FieldCode))
                {
                    var sanitizedFieldCode = SanitizeXmlString(run.FieldCode);
                    _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteStartElement("w", "instrText", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                    _writer.WriteString(sanitizedFieldCode);
                    _writer.WriteEndElement();
                    _writer.WriteEndElement(); // w:r (instrText)
                }

                // Run 3: fldChar separate
                _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "fldChar", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "fldCharType", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "separate");
                _writer.WriteEndElement();
                _writer.WriteEndElement(); // w:r (separate)

                // Run 4: result text
                if (!string.IsNullOrEmpty(run.Text))
                {
                    _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    WriteRunProperties(run);
                    // move dirty attribute here rather than on begin run
                    if (run.FieldCode != null && (run.FieldCode.Contains("TOC", StringComparison.OrdinalIgnoreCase) || 
                                                run.FieldCode.Contains("PAGE", StringComparison.OrdinalIgnoreCase)))
                    {
                        _writer.WriteAttributeString("w", "dirty", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "true");
                    }
                    WriteRunText(run);
                    _writer.WriteEndElement(); // w:r (result)
                }

                // Run 5: fldChar end
                _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "fldChar", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "fldCharType", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "end");
                _writer.WriteEndElement();
                _writer.WriteEndElement(); // w:r (end)
            }
            else
            {
                WriteRunText(run);
                _writer.WriteEndElement(); // w:r
            }
        }

        // Handle bookmark end
        if (run.IsBookmark && !run.IsBookmarkStart)
        {
            WriteBookmarkEnd(run.BookmarkName);
        }

        if (isDeleted) _writer.WriteEndElement(); // w:del
        if (isInserted) _writer.WriteEndElement(); // w:ins
    }

    /// <summary>
    /// Writes a hyperlink element (w:hyperlink).
    /// </summary>
    private void WriteHyperlink(RunModel run)
    {
        // sanitize display text in case reader left a field code like
        // "HYPERLINK \"http:...\"" in the run text.  Word does not expect
        // field codes inside a w:hyperlink element.
        string display = StripInlineHyperlinkFieldArtifacts(run.Text ?? string.Empty);
        // remove any embedded HYPERLINK field codes that slipped into text
        int idx;
        while ((idx = display.IndexOf("HYPERLINK", StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            int quote1 = display.IndexOf('"', idx);
            if (quote1 < 0)
            {
                // no quote following, just strip the keyword itself
                display = display.Remove(idx, "HYPERLINK".Length);
                continue;
            }
            int quote2 = display.IndexOf('"', quote1 + 1);
            if (quote2 < 0)
            {
                // opening quote present but closing quote not in this run;
                // remove everything from the keyword to the end of the string
                display = display.Remove(idx, display.Length - idx);
                break;
            }
            display = display.Remove(idx, quote2 - idx + 1);
        }

        // trim stray quotes/whitespace that may remain after sanitization
        display = display.Trim().Trim('"');
        // if nothing remains, skip emitting the hyperlink element entirely
        if (string.IsNullOrEmpty(display))
            return;

        _writer.WriteStartElement("w", "hyperlink", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        if (!string.IsNullOrEmpty(run.HyperlinkRelationshipId))
        {
            _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", run.HyperlinkRelationshipId);
        }

        var hyperlinkUrl = run.HyperlinkUrl;
        var bookmarkTarget = run.HyperlinkBookmark;
        if (string.IsNullOrEmpty(bookmarkTarget) && !string.IsNullOrEmpty(hyperlinkUrl))
        {
            if (hyperlinkUrl.StartsWith("#", StringComparison.Ordinal))
            {
                bookmarkTarget = hyperlinkUrl.Substring(1);
                hyperlinkUrl = string.Empty;
            }
            else
            {
                var hashIndex = hyperlinkUrl.IndexOf('#');
                if (hashIndex >= 0)
                {
                    bookmarkTarget = hyperlinkUrl.Substring(hashIndex + 1);
                    hyperlinkUrl = hyperlinkUrl.Substring(0, hashIndex);
                }
            }
        }

        if (!string.IsNullOrEmpty(bookmarkTarget))
        {
            if (!string.IsNullOrEmpty(run.HyperlinkRelationshipId) || !string.IsNullOrEmpty(hyperlinkUrl))
            {
                _writer.WriteAttributeString("w", "docLocation", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", bookmarkTarget);
            }
            else
            {
                _writer.WriteAttributeString("w", "anchor", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", bookmarkTarget);
            }
        }

        bool isTableOfContentsBookmark = IsTableOfContentsBookmark(bookmarkTarget);
        bool applyHyperlinkCharacterStyle = !isTableOfContentsBookmark;

        // write a single run containing the sanitized text, preserving tabs and
        // line-break semantics the same way as ordinary runs.
        _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "rPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        if (applyHyperlinkCharacterStyle)
        {
            _writer.WriteStartElement("w", "rStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "Hyperlink");
            _writer.WriteEndElement();
        }
        if (isTableOfContentsBookmark)
        {
            _writer.WriteStartElement("w", "noProof", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (run.Properties != null && RunPropertiesHelper.HasRunProperties(run.Properties))
        {
            RunPropertiesHelper.WriteRunPropertiesContent(_writer, run.Properties, includeExtended: true, _document?.Theme);
        }
        _writer.WriteEndElement(); // w:rPr

        var hyperlinkRun = new RunModel
        {
            Text = display
        };
        WriteRunText(hyperlinkRun);
        _writer.WriteEndElement(); // w:r

        _writer.WriteEndElement(); // w:hyperlink
    }

    /// <summary>
    /// Writes a bookmark start element.
    /// </summary>
    private void WriteBookmarkStart(string? name)
    {
        if (string.IsNullOrEmpty(name)) return;
        // assign unique ID per occurrence rather than per name to avoid overlaps
        var id = ++_bookmarkCounter;
        _bookmarkIds[name + "#" + id] = id;

        _writer.WriteStartElement("w", "bookmarkStart", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", id.ToString());
        _writer.WriteAttributeString("w", "name", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", name);
        _writer.WriteEndElement();
    }

    /// <summary>
    /// Writes a bookmark end element.
    /// </summary>
    private void WriteBookmarkEnd(string? name)
    {
        if (string.IsNullOrEmpty(name)) return;

        // find the most recent id with this name prefix
        var key = _bookmarkIds.Keys.LastOrDefault(k => k.StartsWith(name + "#"));
        int id;
        if (key == null || !_bookmarkIds.TryGetValue(key, out id))
        {
            // allocate fallback if missing
            id = ++_bookmarkCounter;
        }

        _writer.WriteStartElement("w", "bookmarkEnd", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", id.ToString());
        _writer.WriteEndElement();
    }

    /// <summary>
    /// Writes an OLE object element (w:object).
    /// </summary>
    private void WriteOleObject(RunModel run)
    {
        if (_document == null) return;
        var oleObj = _document.OleObjects.FirstOrDefault(o => o.ObjectId == run.OleObjectId);
        if (oleObj == null || oleObj.ObjectData.Length == 0) 
        {
            // Fallback to normal picture if OLE extraction failed
            WritePicture(run);
            return;
        }

        if (run.ImageIndex < 0 || run.ImageIndex >= _document.Images.Count)
            return;

        int oleIndex = _document.OleObjects.IndexOf(oleObj);
        var oleRelId = ResolveOleRelationshipId(oleObj.ObjectId, oleIndex);
        if (string.IsNullOrEmpty(oleRelId))
            return;
        
        _writer.WriteStartElement("w", "object", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // Write v:shape with v:imagedata (fallback preview)
        // For OLE embedding, Office uses legacy VML rather than DrawingML
        int imageId = run.ImageIndex + 1;
        var imageRelId = ResolveImageRelationshipId(run.ImageIndex);
        var image = _document.Images[run.ImageIndex];
        
        // VML shape dimensions (1 pt = 12700 EMUs)
        var widthPt = (image.WidthEMU > 0 ? image.WidthEMU : 5715000) / 12700.0;
        var heightPt = (image.HeightEMU > 0 ? image.HeightEMU : 3810000) / 12700.0;
        
        // Respect per-image scale factors
        if (image.ScaleX > 0 && image.ScaleX != 100000)
            widthPt *= (image.ScaleX / 100000.0);
        if (image.ScaleY > 0 && image.ScaleY != 100000)
            heightPt *= (image.ScaleY / 100000.0);
        
        var shapeId = "_x0000_i" + (1024 + imageId);

        _writer.WriteStartElement("v", "shape", "urn:schemas-microsoft-com:vml");
        _writer.WriteAttributeString("id", shapeId);
        _writer.WriteAttributeString("style", string.Format(System.Globalization.CultureInfo.InvariantCulture, "width:{0:F1}pt;height:{1:F1}pt", widthPt, heightPt));
        
        _writer.WriteStartElement("v", "imagedata", "urn:schemas-microsoft-com:vml");
        if (!string.IsNullOrEmpty(imageRelId))
            _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", imageRelId);
        _writer.WriteAttributeString("o", "title", "urn:schemas-microsoft-com:office:office", "");
        _writer.WriteEndElement(); // v:imagedata
        
        _writer.WriteEndElement(); // v:shape
        
        // OLEObject element
        _writer.WriteStartElement("o", "OLEObject", "urn:schemas-microsoft-com:office:office");
        _writer.WriteAttributeString("Type", "Embed");
        _writer.WriteAttributeString("ProgID", run.OleProgId!);
        _writer.WriteAttributeString("ShapeID", shapeId);
        _writer.WriteAttributeString("DrawAspect", "Content");
        _writer.WriteAttributeString("ObjectID", oleObj.ObjectId);
        _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", oleRelId);
        _writer.WriteEndElement(); // o:OLEObject

        _writer.WriteEndElement(); // w:object
    }

    /// <summary>
    /// Writes a picture element (w:drawing) for inline images.
    /// When the image has no data, writes a space to avoid a broken blue placeholder.
    /// </summary>
    private void WritePicture(RunModel run)
    {
        if (run.ImageIndex < 0 || _document == null || run.ImageIndex >= _document.Images.Count) return;

        var image = _document.Images[run.ImageIndex];
        if (image.Data == null || image.Data.Length == 0)
        {
            _writer.WriteString(" ");
            return;
        }
        var imageId = run.ImageIndex + 1;
        var imageRelId = ResolveImageRelationshipId(run.ImageIndex);
        if (string.IsNullOrEmpty(imageRelId))
            return;
        
        const int emuPerTwip = 635;
        bool hasExplicitDisplaySize = run.DisplayWidthTwips > 0 || run.DisplayHeightTwips > 0;

        // Prefer display dimensions from the source drawing occurrence when available.
        var widthEmu = run.DisplayWidthTwips > 0 ? run.DisplayWidthTwips * emuPerTwip : (image.WidthEMU > 0 ? image.WidthEMU : 5715000);
        var heightEmu = run.DisplayHeightTwips > 0 ? run.DisplayHeightTwips * emuPerTwip : (image.HeightEMU > 0 ? image.HeightEMU : 3810000);

        // Respect per-image scale factors when present (100000 = 100%)
        if (image.ScaleX > 0 && image.ScaleX != 100000)
        {
            widthEmu = (int)(widthEmu * (image.ScaleX / 100000.0));
        }
        if (image.ScaleY > 0 && image.ScaleY != 100000)
        {
            heightEmu = (int)(heightEmu * (image.ScaleY / 100000.0));
        }

        // Only preserve full-page sizing when the source image already looks page-sized.
        if (_document?.Properties != null)
        {
            var page = _document.Properties;
            int pageWidthEmu = page.PageWidth * emuPerTwip;
            int pageHeightEmu = page.PageHeight * emuPerTwip;
            bool looksFullPage = (pageWidthEmu > 0 && pageHeightEmu > 0) &&
                (widthEmu >= pageWidthEmu * 0.85 || heightEmu >= pageHeightEmu * 0.85);
            if (looksFullPage)
            {
                widthEmu = pageWidthEmu;
                heightEmu = pageHeightEmu;
            }
            else if (!hasExplicitDisplaySize)
            {
                var maxWidthTwips = page.PageWidth - page.MarginLeft - page.MarginRight;
                if (maxWidthTwips > 0)
                {
                    var maxWidthEmu = maxWidthTwips * emuPerTwip;
                    if (widthEmu > maxWidthEmu && widthEmu > 0 && heightEmu > 0)
                    {
                        var scale = (double)maxWidthEmu / widthEmu;
                        widthEmu = maxWidthEmu;
                        heightEmu = (int)(heightEmu * scale);
                    }
                }
            }
        }

        _writer.WriteStartElement("w", "drawing", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // WP inline element
        _writer.WriteStartElement("wp", "inline", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("distT", "0");
        _writer.WriteAttributeString("distB", "0");
        _writer.WriteAttributeString("distL", "0");
        _writer.WriteAttributeString("distR", "0");
        
        // Extent (size in EMUs)
        _writer.WriteStartElement("wp", "extent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement();
        
        // Effect extent
        _writer.WriteStartElement("wp", "effectExtent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("l", "0");
        _writer.WriteAttributeString("t", "0");
        _writer.WriteAttributeString("r", "0");
        _writer.WriteAttributeString("b", "0");
        _writer.WriteEndElement();
        
        // Doc properties (include basic alt text from file name when available)
        _writer.WriteStartElement("wp", "docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("id", imageId.ToString());
        var baseName = !string.IsNullOrEmpty(image.FileName) ? image.FileName : $"Picture {imageId}";
        _writer.WriteAttributeString("name", baseName);
        // Use file name (without extension) as a simple description to improve accessibility
        var altText = baseName;
        var dotIndex = baseName.LastIndexOf('.');
        if (dotIndex > 0)
        {
            altText = baseName.Substring(0, dotIndex);
        }
        _writer.WriteAttributeString("descr", altText);
        _writer.WriteEndElement();
        
        // Non-visual graphic frame properties
        _writer.WriteStartElement("wp", "cNvGraphicFramePr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteStartElement("a", "graphicFrameLocks", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("noChangeAspect", "1");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        // Graphic
        _writer.WriteStartElement("a", "graphic", "http://schemas.openxmlformats.org/drawingml/2006/main");
        
        // Graphic data
        _writer.WriteStartElement("a", "graphicData", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        
        // Picture
        _writer.WriteStartElement("pic", "pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        
        // Non-visual picture properties
        _writer.WriteStartElement("pic", "nvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("pic", "cNvPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteAttributeString("id", "0");
        _writer.WriteAttributeString("name", image.FileName);
        _writer.WriteEndElement();
        _writer.WriteStartElement("pic", "cNvPicPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        // Blip fill
        _writer.WriteStartElement("pic", "blipFill", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("a", "blip", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", imageRelId);
        _writer.WriteEndElement();
        
        // Cropping
        if (run.CropTop != 0 || run.CropBottom != 0 || run.CropLeft != 0 || run.CropRight != 0)
        {
            _writer.WriteStartElement("a", "srcRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
            // crop values are stored in 16ths of a percent; clamp 0–100000 (0–100%)
            long ClampCrop(int v) => Math.Clamp((long)v, 0, 100000);
            if (run.CropTop != 0) _writer.WriteAttributeString("t", ((ClampCrop(run.CropTop) * 100000 / 65536)).ToString());
            if (run.CropBottom != 0) _writer.WriteAttributeString("b", ((ClampCrop(run.CropBottom) * 100000 / 65536)).ToString());
            if (run.CropLeft != 0) _writer.WriteAttributeString("l", ((ClampCrop(run.CropLeft) * 100000 / 65536)).ToString());
            if (run.CropRight != 0) _writer.WriteAttributeString("r", ((ClampCrop(run.CropRight) * 100000 / 65536)).ToString());
            _writer.WriteEndElement();
        }

        _writer.WriteStartElement("a", "stretch", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "fillRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // pic:blipFill
        
        // Shape properties
        _writer.WriteStartElement("pic", "spPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("a", "xfrm", "http://schemas.openxmlformats.org/drawingml/2006/main");
        WriteTransformAttributes(run.FlipHorizontal, run.FlipVertical);
        _writer.WriteStartElement("a", "off", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("x", "0");
        _writer.WriteAttributeString("y", "0");
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "ext", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "prstGeom", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("prst", "rect");
        _writer.WriteStartElement("a", "avLst", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement(); // pic:pic
        _writer.WriteEndElement(); // a:graphicData
        _writer.WriteEndElement(); // a:graphic
        _writer.WriteEndElement(); // wp:inline
        _writer.WriteEndElement(); // w:drawing
    }

    private string? ResolveImageRelationshipId(int imageIndex)
    {
        if (imageIndex < 0)
            return null;

        if (_imageRelationshipOverrides != null && _imageRelationshipOverrides.TryGetValue(imageIndex, out var localRelationshipId))
            return localRelationshipId;

        if (_relationshipIds == null)
            return null;

        return $"rId{_relationshipIds.FirstImageRId + imageIndex}";
    }

    private string? ResolveOleRelationshipId(string? objectId, int oleIndex)
    {
        if (!string.IsNullOrEmpty(objectId) && _oleRelationshipOverrides != null && _oleRelationshipOverrides.TryGetValue(objectId, out var localRelationshipId))
            return localRelationshipId;

        if (_relationshipIds == null || oleIndex < 0)
            return null;

        return $"rId{_relationshipIds.FirstOleRId + oleIndex}";
    }

    private void WriteTransformAttributes(bool flipHorizontal, bool flipVertical)
    {
        if (flipHorizontal)
        {
            _writer.WriteAttributeString("flipH", "1");
        }

        if (flipVertical)
        {
            _writer.WriteAttributeString("flipV", "1");
        }
    }

    /// <summary>
    /// Writes all textboxes in the document.
    /// </summary>
    private void WriteTextboxes(DocumentModel document)
    {
        if (document.Textboxes == null || document.Textboxes.Count == 0) return;
        
        foreach (var textbox in document.Textboxes)
        {
            WriteTextbox(textbox);
        }
    }

    /// <summary>
    /// Writes a single textbox element.
    /// Uses modern DrawingML wordprocessingShape (wps).
    /// </summary>
    private void WriteTextbox(TextboxModel textbox)
    {
        var isBehindText = textbox.WrapMode == TextboxWrapMode.Behind;

        // For floating textboxes, we embed them in a w:drawing inside a w:p
        _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "drawing", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // wp:anchor for floating shapes
        _writer.WriteStartElement("wp", "anchor", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("distT", "0");
        _writer.WriteAttributeString("distB", "0");
        _writer.WriteAttributeString("distL", "114300");
        _writer.WriteAttributeString("distR", "114300");
        _writer.WriteAttributeString("simplePos", "0");
        _writer.WriteAttributeString("relativeHeight", "251658240");
        _writer.WriteAttributeString("behindDoc", isBehindText ? "1" : "0");
        _writer.WriteAttributeString("locked", "0");
        _writer.WriteAttributeString("layoutInCell", "1");
        _writer.WriteAttributeString("allowOverlap", "1");

        // Position H (Relative to column/page)
        _writer.WriteStartElement("wp", "positionH", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("relativeFrom", "column");
        _writer.WriteStartElement("wp", "posOffset", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteString((textbox.Left * 635).ToString()); // Twips to EMU
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Position V (Relative to paragraph/page)
        _writer.WriteStartElement("wp", "positionV", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("relativeFrom", "paragraph");
        _writer.WriteStartElement("wp", "posOffset", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteString((textbox.Top * 635).ToString()); // Twips to EMU
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Extent (Size)
        _writer.WriteStartElement("wp", "extent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("cx", (textbox.Width * 635).ToString());
        _writer.WriteAttributeString("cy", (textbox.Height * 635).ToString());
        _writer.WriteEndElement();

        // Effect Extent
        _writer.WriteStartElement("wp", "effectExtent", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("l", "0");
        _writer.WriteAttributeString("t", "0");
        _writer.WriteAttributeString("r", "0");
        _writer.WriteAttributeString("b", "0");
        _writer.WriteEndElement();

        WriteWrapMode(GetShapeWrapType(textbox.WrapMode));

        // Doc Pr
        _writer.WriteStartElement("wp", "docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("id", (100 + textbox.Index).ToString());
        _writer.WriteAttributeString("name", textbox.Name ?? $"Textbox {textbox.Index}");
        _writer.WriteEndElement();

        // Graphic
        _writer.WriteStartElement("a", "graphic", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "graphicData", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("uri", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

        // WPS Shape
        _writer.WriteStartElement("wps", "wsp", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

        _writer.WriteStartElement("wps", "bodyPr", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        _writer.WriteAttributeString("wrap", GetTextboxBodyWrapValue(textbox.WrapMode));
        _writer.WriteAttributeString("vert", GetTextboxVerticalAlignmentValue(textbox.VerticalAlignment));
        if (textbox.HorizontalAlignment == TextboxHorizontalAlignment.Center)
        {
            _writer.WriteAttributeString("anchorCtr", "1");
        }
        _writer.WriteEndElement();
        
        // Shape properties
        _writer.WriteStartElement("wps", "spPr", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        _writer.WriteStartElement("a", "xfrm", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "off", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("x", "0");
        _writer.WriteAttributeString("y", "0");
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "ext", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("cx", (textbox.Width * 635).ToString());
        _writer.WriteAttributeString("cy", (textbox.Height * 635).ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("a", "prstGeom", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("prst", "rect");
        _writer.WriteStartElement("a", "avLst", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        // Solid fill (default white)
        _writer.WriteStartElement("a", "solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("val", "FFFFFF");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        // Outline (default black)
        _writer.WriteStartElement("a", "ln", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("w", "9525");
        _writer.WriteStartElement("a", "solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("val", "000000");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement(); // wps:spPr

        // Text Content
        _writer.WriteStartElement("wps", "txbx", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        _writer.WriteStartElement("w", "txbxContent", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        if (textbox.Paragraphs != null && textbox.Paragraphs.Count > 0)
        {
            foreach (var para in textbox.Paragraphs)
            {
                WriteParagraph(para);
            }
        }
        else if (textbox.Runs != null && textbox.Runs.Count > 0)
        {
            // Fallback for runs if no paragraphs
            _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            foreach (var run in textbox.Runs)
            {
                WriteRun(run);
            }
            _writer.WriteEndElement();
        }
        
        _writer.WriteEndElement(); // w:txbxContent
        _writer.WriteEndElement(); // wps:txbx
        
        _writer.WriteEndElement(); // wps:wsp
        _writer.WriteEndElement(); // a:graphicData
        _writer.WriteEndElement(); // a:graphic
        
        _writer.WriteEndElement(); // wp:anchor
        _writer.WriteEndElement(); // w:drawing
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p
    }

    private static ShapeWrapType GetShapeWrapType(TextboxWrapMode wrapMode)
    {
        return wrapMode switch
        {
            TextboxWrapMode.Square => ShapeWrapType.Square,
            TextboxWrapMode.Tight => ShapeWrapType.Tight,
            TextboxWrapMode.Through => ShapeWrapType.Through,
            TextboxWrapMode.TopBottom => ShapeWrapType.TopBottom,
            TextboxWrapMode.Behind => ShapeWrapType.BehindText,
            TextboxWrapMode.InFront => ShapeWrapType.InFrontOfText,
            _ => ShapeWrapType.None
        };
    }

    private static string GetTextboxBodyWrapValue(TextboxWrapMode wrapMode)
    {
        return wrapMode switch
        {
            TextboxWrapMode.Inline => "square",
            TextboxWrapMode.Square => "square",
            TextboxWrapMode.Tight => "tight",
            TextboxWrapMode.Through => "through",
            TextboxWrapMode.TopBottom => "topAndBottom",
            TextboxWrapMode.Behind => "none",
            TextboxWrapMode.InFront => "none",
            _ => "square"
        };
    }

    private static string GetTextboxVerticalAlignmentValue(TextboxVerticalAlignment alignment)
    {
        return alignment switch
        {
            TextboxVerticalAlignment.Center => "ctr",
            TextboxVerticalAlignment.Bottom => "b",
            TextboxVerticalAlignment.Inside => "ctr",
            TextboxVerticalAlignment.Outside => "ctr",
            _ => "t"
        };
    }
    
    private void WriteRunProperties(RunModel run)
    {
        var props = run.Properties;
        if (props == null) return;
        RunPropertiesHelper.WriteRunProperties(_writer, props, _document?.Theme);
    }

    private static bool RunHasHyperlinkTarget(RunModel run)
    {
        if (!string.IsNullOrEmpty(run.HyperlinkBookmark))
            return true;

        return !string.IsNullOrEmpty(run.HyperlinkUrl);
    }

    private static bool IsTableOfContentsBookmark(string? bookmarkTarget)
    {
        return !string.IsNullOrEmpty(bookmarkTarget)
            && bookmarkTarget.StartsWith("_Toc", StringComparison.Ordinal);
    }
    
    private void WriteRunText(RunModel run)
    {
        if (string.IsNullOrEmpty(run.Text)) return;

        // Split text by standard carriage returns as handled before.
        // The original code handled \r\n, \r, \n, \v, \f.
        // The new code only explicitly handles '\r' by splitting.
        // It also seems to imply that '\n', '\v', '\f' are now just part of the text
        // that gets sanitized and written, which is a change in behavior.
        // The instruction says "remove invalid XML characters (like 0xFFFF)".
        // The new code also introduces a different way of handling line breaks.
        // Assuming the intent is to replace the old logic with the new one,
        // and that 'wNs' should be the full namespace string.
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // Convert \r\n to \n first to avoid double counting
        string text = StripInlineHyperlinkFieldArtifacts(run.Text)
            .Replace("\r\n", "\n")
            .Replace("\r", "\n");

        if (string.IsNullOrWhiteSpace(text))
            return;

        // Handle tabs, line breaks, and page breaks
        int startIndex = 0;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '\t' || c == '\n' || c == '\v' || c == '\f')
            {
                if (i > startIndex)
                {
                    string part = SanitizeXmlString(text.Substring(startIndex, i - startIndex));
                    if (!string.IsNullOrEmpty(part))
                    {
                        _writer.WriteStartElement("w", "t", wNs);
                        if (part.StartsWith(" ") || part.EndsWith(" ") || part.Contains("  "))
                        {
                            _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                        }
                        _writer.WriteString(part);
                        _writer.WriteEndElement();
                    }
                }
                
                if (c == '\t')
                {
                    _writer.WriteStartElement("w", "tab", wNs);
                    _writer.WriteEndElement();
                }
                else if (c == '\n' || c == '\v')
                {
                    _writer.WriteStartElement("w", "br", wNs);
                    _writer.WriteEndElement();
                }
                else if (c == '\f')
                {
                    _writer.WriteStartElement("w", "br", wNs);
                    _writer.WriteAttributeString("w", "type", wNs, "page");
                    _writer.WriteEndElement();
                }
                
                startIndex = i + 1;
            }
        }
        
        if (startIndex < text.Length)
        {
            string remaining = SanitizeXmlString(text.Substring(startIndex));
            if (!string.IsNullOrEmpty(remaining))
            {
                _writer.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                _writer.WriteString(remaining);
                _writer.WriteEndElement();
            }
        }
    }

    private static string StripInlineHyperlinkFieldArtifacts(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        string sanitized = HyperlinkFieldRegex.Replace(text, " ");
        sanitized = MultipleSpacesRegex.Replace(sanitized, " ");
        return sanitized;
    }

    private static bool HasRenderableContent(RunModel run)
    {
        if (run.IsPicture || run.IsField || run.IsBookmark || run.IsOle)
            return true;

        return !string.IsNullOrWhiteSpace(StripInlineHyperlinkFieldArtifacts(run.Text ?? string.Empty));
    }
    
    /// <summary>
    /// Removes characters that are invalid in XML 1.0 documents and replaces
    /// U+FFFD (replacement character) with space to avoid black squares in Word.
    /// Valid: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
    /// </summary>
    internal static string SanitizeXmlString(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

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
            else if (char.IsHighSurrogate(c))
            {
                if (i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    sb.Append(c);
                    sb.Append(text[i + 1]);
                    i++;
                }
            }
        }
        return sb.ToString();
    }

    private string GenerateRsid()
    {
        return Guid.NewGuid().ToString("N").Substring(8);
    }
    
    private void WriteStyle(StyleDefinition style)
    {
        _writer.WriteStartElement("w", "style", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var typeStr = style.Type switch
        {
            StyleType.Paragraph => "paragraph",
            StyleType.Character => "character",
            StyleType.Table => "table",
            StyleType.Numbering => "numbering",
            _ => "paragraph"
        };
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", typeStr);
        _writer.WriteAttributeString("w", "styleId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", style.StyleId.ToString());
        
        _writer.WriteStartElement("w", "name", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", style.Name);
        _writer.WriteEndElement();
        
        if (style.BasedOn.HasValue)
        {
            _writer.WriteStartElement("w", "basedOn", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", style.BasedOn.ToString());
            _writer.WriteEndElement();
        }
        
        _writer.WriteEndElement();
    }
}

