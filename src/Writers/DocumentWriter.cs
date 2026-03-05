using System.Xml;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Writes DOCX document using XmlWriter for optimal streaming performance
/// </summary>
public class DocumentWriter
{
    private readonly XmlWriter _writer;
    private int _runId = 0;
    private int _trackChangeId = 1;
    private int _trackChangeId = 1;
    private DocumentModel? _document;
    private DocumentRelationshipIds? _relationshipIds;
    private readonly Dictionary<string, int> _bookmarkIds = new(StringComparer.Ordinal);
    private int _bookmarkCounter = 0;
    private System.Collections.Generic.HashSet<string> _startedComments = new();
    private System.Collections.Generic.HashSet<string> _endedComments = new();
    private HashSet<string> _startedComments = new();
    private HashSet<string> _endedComments = new();
    /// <summary>When true, do not emit pageBreakBefore so leading content (e.g. 绿色等级评价报告) stays on page 1.</summary>
    private bool _suppressLeadingPageBreak;
    /// <summary>When true, the next picture written in the body should use full-page dimensions (first-page background).</summary>
    private bool _firstBodyPictureNotYetWritten;

    public DocumentWriter(XmlWriter writer)
    {
        _writer = writer;
    }
    
    /// <summary>
    /// Writes a basic vector shape (non-picture OfficeArt / SmartArt fallback)
    /// as a DrawingML wordprocessingShape rectangle. Position and size are
    /// taken from ShapeAnchor when available; otherwise a reasonable inline
    /// size is used.
    /// </summary>
    private void WriteVectorShape(ShapeModel shape)
    {
        if (_document == null)
            return;

        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string wpsNs = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        // Derive size in EMUs from anchor or use a default (~4x3 inches).
        const int emuPerTwip = 635;
        int widthEmu = 3657600;  // 4"
        int heightEmu = 2743200; // 3"
        int xEmu = 0;
        int yEmu = 0;
        bool isFloating = false;

        if (shape.Anchor != null)
        {
            var anchor = shape.Anchor;
            if (anchor.Width > 0)
            {
                widthEmu = anchor.Width * emuPerTwip;
            }
            if (anchor.Height > 0)
            {
                heightEmu = anchor.Height * emuPerTwip;
            }
            xEmu = Math.Max(0, anchor.X * emuPerTwip);
            yEmu = Math.Max(0, anchor.Y * emuPerTwip);
            isFloating = anchor.IsFloating;
        }

        // Clamp width to page width (inside margins) while preserving aspect ratio.
        if (_document.Properties != null)
        {
            var page = _document.Properties;
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

        _writer.WriteStartElement("w", "p", wNs);
        _writer.WriteStartElement("w", "r", wNs);
        _writer.WriteStartElement("w", "drawing", wNs);

        if (isFloating)
        {
            // Floating vector shape using wp:anchor (similar to floating picture).
            _writer.WriteStartElement("wp", "anchor", wpNs);
            _writer.WriteAttributeString("distT", "0");
            _writer.WriteAttributeString("distB", "0");
            _writer.WriteAttributeString("distL", "0");
            _writer.WriteAttributeString("distR", "0");
            _writer.WriteAttributeString("simplePos", "0");
            var relHeight = shape.Anchor?.ZOrder ?? 0;
            if (relHeight < 0) relHeight = 0;
            _writer.WriteAttributeString("relativeHeight", relHeight.ToString());
            _writer.WriteAttributeString("behindDoc", "0");
            _writer.WriteAttributeString("locked", "0");
            _writer.WriteAttributeString("layoutInCell", "1");
            _writer.WriteAttributeString("allowOverlap", "1");

            // Horizontal & vertical position.
            _writer.WriteStartElement("wp", "positionH", wpNs);
            _writer.WriteAttributeString("relativeFrom", "page");
            _writer.WriteStartElement("wp", "posOffset", wpNs);
            _writer.WriteString(xEmu.ToString());
            _writer.WriteEndElement(); // wp:posOffset
            _writer.WriteEndElement(); // wp:positionH

            _writer.WriteStartElement("wp", "positionV", wpNs);
            _writer.WriteAttributeString("relativeFrom", "page");
            _writer.WriteStartElement("wp", "posOffset", wpNs);
            _writer.WriteString(yEmu.ToString());
            _writer.WriteEndElement(); // wp:posOffset
            _writer.WriteEndElement(); // wp:positionV

            // Extent
            _writer.WriteStartElement("wp", "extent", wpNs);
            _writer.WriteAttributeString("cx", widthEmu.ToString());
            _writer.WriteAttributeString("cy", heightEmu.ToString());
            _writer.WriteEndElement();

            // Effect extent
            _writer.WriteStartElement("wp", "effectExtent", wpNs);
            _writer.WriteAttributeString("l", "0");
            _writer.WriteAttributeString("t", "0");
            _writer.WriteAttributeString("r", "0");
            _writer.WriteAttributeString("b", "0");
            _writer.WriteEndElement();

            // Text wrapping: square by default for shapes.
            _writer.WriteStartElement("wp", "wrapSquare", wpNs);
            _writer.WriteAttributeString("wrapText", "bothSides");
            _writer.WriteEndElement();

            // docPr
            _writer.WriteStartElement("wp", "docPr", wpNs);
            _writer.WriteAttributeString("id", (2000 + shape.Id).ToString());
            _writer.WriteAttributeString("name", $"Shape {shape.Id}");
            _writer.WriteEndElement(); // wp:docPr

            // Graphic frame props
            _writer.WriteStartElement("wp", "cNvGraphicFramePr", wpNs);
            _writer.WriteStartElement("a", "graphicFrameLocks", aNs);
            _writer.WriteAttributeString("noChangeAspect", "1");
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            // Graphic
            _writer.WriteStartElement("a", "graphic", aNs);
            _writer.WriteStartElement("a", "graphicData", aNs);
            _writer.WriteAttributeString("uri", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            // WordprocessingShape (rectangle/ellipse etc.)
            WriteWpsShape(shape, widthEmu, heightEmu);

            _writer.WriteEndElement(); // a:graphicData
            _writer.WriteEndElement(); // a:graphic

            _writer.WriteEndElement(); // wp:anchor
        }
        else
        {
            // Inline vector shape using wp:inline.
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

            _writer.WriteStartElement("wp", "docPr", wpNs);
            _writer.WriteAttributeString("id", (2000 + shape.Id).ToString());
            _writer.WriteAttributeString("name", $"Shape {shape.Id}");
            _writer.WriteEndElement(); // wp:docPr

            _writer.WriteStartElement("wp", "cNvGraphicFramePr", wpNs);
            _writer.WriteStartElement("a", "graphicFrameLocks", aNs);
            _writer.WriteAttributeString("noChangeAspect", "1");
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "graphic", aNs);
            _writer.WriteStartElement("a", "graphicData", aNs);
            _writer.WriteAttributeString("uri", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            WriteWpsShape(shape, widthEmu, heightEmu);

            _writer.WriteEndElement(); // a:graphicData
            _writer.WriteEndElement(); // a:graphic

            _writer.WriteEndElement(); // wp:inline
        }

        _writer.WriteEndElement(); // w:drawing
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p
    }

    /// <summary>
    /// Writes the inner wps:wsp contents (geometry and basic styling) for a
    /// vector shape.
    /// </summary>
    private void WriteWpsShape(ShapeModel shape, int widthEmu, int heightEmu)
    {
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string wpsNs = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        _writer.WriteStartElement("wps", "wsp", wpsNs);

        // spPr
        _writer.WriteStartElement("wps", "spPr", wpsNs);
        _writer.WriteStartElement("a", "xfrm", aNs);
        _writer.WriteStartElement("a", "off", aNs);
        _writer.WriteAttributeString("x", "0");
        _writer.WriteAttributeString("y", "0");
        _writer.WriteEndElement(); // a:off
        _writer.WriteStartElement("a", "ext", aNs);
        _writer.WriteAttributeString("cx", widthEmu.ToString());
        _writer.WriteAttributeString("cy", heightEmu.ToString());
        _writer.WriteEndElement(); // a:ext
        _writer.WriteEndElement(); // a:xfrm

        // Geometry: rectangle or ellipse; default rectangle.
        string prst = shape.Type switch
        {
            ShapeType.Ellipse => "ellipse",
            _ => "rect"
        };
        _writer.WriteStartElement("a", "prstGeom", aNs);
        _writer.WriteAttributeString("prst", prst);
        _writer.WriteStartElement("a", "avLst", aNs);
        _writer.WriteEndElement(); // a:avLst
        _writer.WriteEndElement(); // a:prstGeom

        // Fill color (if any)
        if (shape.FillColor != 0)
        {
            _writer.WriteStartElement("a", "solidFill", aNs);
            _writer.WriteStartElement("a", "srgbClr", aNs);
            var fillHex = ColorHelper.ColorToHex(shape.FillColor);
            if (fillHex == "auto") fillHex = "FFFFFF";
            _writer.WriteAttributeString("val", fillHex);
            _writer.WriteEndElement(); // a:srgbClr
            _writer.WriteEndElement(); // a:solidFill
        }

        // Line (stroke)
        _writer.WriteStartElement("a", "ln", aNs);
        if (shape.LineWidth > 0)
        {
            _writer.WriteAttributeString("w", shape.LineWidth.ToString());
        }
        _writer.WriteStartElement("a", "solidFill", aNs);
        _writer.WriteStartElement("a", "srgbClr", aNs);
        var lineHex = shape.LineColor != 0 ? ColorHelper.ColorToHex(shape.LineColor) : "000000";
        if (lineHex == "auto") lineHex = "000000";
        _writer.WriteAttributeString("val", lineHex);
        _writer.WriteEndElement(); // a:srgbClr
        _writer.WriteEndElement(); // a:solidFill
        _writer.WriteEndElement(); // a:ln

        _writer.WriteEndElement(); // wps:spPr

        // Optionally, basic textbox for shape text when available.
        if (!string.IsNullOrEmpty(shape.Text))
        {
            _writer.WriteStartElement("wps", "txbx", wpsNs);
            _writer.WriteStartElement("w", "txbxContent", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
            _writer.WriteString(shape.Text);
            _writer.WriteEndElement(); // w:t
            _writer.WriteEndElement(); // w:r
            _writer.WriteEndElement(); // w:p
            _writer.WriteEndElement(); // w:txbxContent
            _writer.WriteEndElement(); // wps:txbx
        }

        _writer.WriteEndElement(); // wps:wsp
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
        var baseName = !string.IsNullOrEmpty(chart.Title) ? chart.Title : $"Chart {chartIndex + 1}";
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
        _document = document;
        _relationshipIds = RelationshipsWriter.ComputeRelationshipIds(document);
        
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "document", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
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

        // First picture in body is often the first-page background; use full-page size for it
        _firstBodyPictureNotYetWritten = true;
        
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
                    const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    _writer.WriteStartElement("w", "sectPr", wNs);
                    WriteSectionContent(sectionForTable);
                    _writer.WriteEndElement();
                }

                paraIndex = table.EndParagraphIndex + 1;
            }
            else
            {
                var paragraph = document.Paragraphs[paraIndex];
                WriteParagraph(paragraph, _suppressLeadingPageBreak);
                if (_suppressLeadingPageBreak && ParagraphHasVisibleContent(paragraph))
                    _suppressLeadingPageBreak = false;

                // If a section ends at this paragraph, emit sectPr immediately after it
                if (sectionEndMap.TryGetValue(paragraph.Index, out var sectionForParagraph))
                {
                    const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    _writer.WriteStartElement("w", "sectPr", wNs);
                    WriteSectionContent(sectionForParagraph);
                    _writer.WriteEndElement();
                }

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

    /// <summary>
    /// Writes a single shape. Picture shapes are rendered either as true
    /// floating images (wp:anchor) when ShapeAnchor.IsFloating is set, or as
    /// inline pictures using the existing run-based logic. Non-picture shapes
    /// (basic OfficeArt vectors and SmartArt fallbacks) are rendered as
    /// DrawingML wordprocessingShape rectangles with simple fill/line styling.
    /// </summary>
    private void WriteInlinePictureShape(ShapeModel shape, DocumentModel document)
    {
        // Picture-backed shapes
        if (shape.Type == ShapeType.Picture)
        {
            if (shape.ImageIndex is null)
                return;

            var imageIndex = shape.ImageIndex.Value;
            if (imageIndex < 0 || imageIndex >= document.Images.Count)
                return;

            // Prefer floating output when we have a valid anchor.
            if (shape.Anchor is { IsFloating: true })
            {
                WriteFloatingPictureShape(shape, document);
                return;
            }

            // Fallback: inline picture using existing run-based logic.
            var run = new RunModel
            {
                IsPicture = true,
                ImageIndex = imageIndex,
                Properties = new RunProperties()
            };

            _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            WriteRun(run);
            _writer.WriteEndElement(); // w:r
            _writer.WriteEndElement(); // w:p
            return;
        }

        // Non-picture vector shapes (complex OfficeArt, SmartArt fallbacks)
        WriteVectorShape(shape);
    }
    
    /// <summary>
    /// Writes a floating picture using wp:anchor based on ShapeAnchor coordinates.
    /// </summary>
    private void WriteFloatingPictureShape(ShapeModel shape, DocumentModel document)
    {
        if (shape.ImageIndex is null || _document == null)
            return;

        var imageIndex = shape.ImageIndex.Value;
        if (imageIndex < 0 || imageIndex >= document.Images.Count)
            return;

        var image = document.Images[imageIndex];
        if (image.Data == null || image.Data.Length == 0)
            return;

        var anchor = shape.Anchor!;

        // Relationship ID and doc-level image id
        var ids = RelationshipsWriter.ComputeRelationshipIds(_document);
        var relId = $"rId{ids.FirstImageRId + imageIndex}";
        var imageId = imageIndex + 1;

        // Compute size in EMUs, preferring anchor size when available.
        const int emuPerTwip = 635;
        int widthEmu;
        int heightEmu;

        if (anchor.Width > 0 && anchor.Height > 0)
        {
            widthEmu = anchor.Width * emuPerTwip;
            heightEmu = anchor.Height * emuPerTwip;
        }
        else
        {
            widthEmu = image.WidthEMU > 0 ? image.WidthEMU : 5715000;
            heightEmu = image.HeightEMU > 0 ? image.HeightEMU : 3810000;
        }

        // Respect per-image scale factors
        if (image.ScaleX > 0 && image.ScaleX != 100000)
        {
            widthEmu = (int)(widthEmu * (image.ScaleX / 100000.0));
        }
        if (image.ScaleY > 0 && image.ScaleY != 100000)
        {
            heightEmu = (int)(heightEmu * (image.ScaleY / 100000.0));
        }

        // Full-page background: first body picture or anchor/size close to page → full page dimensions
        if (_document.Properties != null)
        {
            var page = _document.Properties;
            int pageWidthEmu = page.PageWidth * emuPerTwip;
            int pageHeightEmu = page.PageHeight * emuPerTwip;
            bool forceFirstFullPage = _firstBodyPictureNotYetWritten && pageWidthEmu > 0 && pageHeightEmu > 0;
            bool looksFullPage = !forceFirstFullPage && (pageWidthEmu > 0 && pageHeightEmu > 0) &&
                (widthEmu >= pageWidthEmu * 0.85 || heightEmu >= pageHeightEmu * 0.85);
            if (forceFirstFullPage || looksFullPage)
            {
                widthEmu = pageWidthEmu;
                heightEmu = pageHeightEmu;
                if (forceFirstFullPage) _firstBodyPictureNotYetWritten = false;
            }
            else
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

        // Convert anchor position from twips to EMUs; clamp to non-negative.
        var xEmu = Math.Max(0, anchor.X * emuPerTwip);
        var yEmu = Math.Max(0, anchor.Y * emuPerTwip);

        _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "drawing", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // Floating anchor
        _writer.WriteStartElement("wp", "anchor", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("distT", "0");
        _writer.WriteAttributeString("distB", "0");
        _writer.WriteAttributeString("distL", "0");
        _writer.WriteAttributeString("distR", "0");
        _writer.WriteAttributeString("simplePos", "0");
        var relHeight = anchor.ZOrder;
        if (relHeight < 0) relHeight = 0;
        _writer.WriteAttributeString("relativeHeight", relHeight.ToString());
        _writer.WriteAttributeString("behindDoc", "0");
        _writer.WriteAttributeString("locked", "0");
        _writer.WriteAttributeString("layoutInCell", "1");
        _writer.WriteAttributeString("allowOverlap", "1");

        // Position
        string MapRelative(ShapeRelativeTo rel) => rel switch
        {
            ShapeRelativeTo.Margin => "margin",
            ShapeRelativeTo.Column => "column",
            ShapeRelativeTo.Paragraph => "paragraph",
            _ => "page"
        };

        _writer.WriteStartElement("wp", "positionH", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("relativeFrom", MapRelative(anchor.HorizontalRelativeTo));
        _writer.WriteStartElement("wp", "posOffset", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteString(xEmu.ToString());
        _writer.WriteEndElement(); // wp:posOffset
        _writer.WriteEndElement(); // wp:positionH

        _writer.WriteStartElement("wp", "positionV", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("relativeFrom", MapRelative(anchor.VerticalRelativeTo));
        _writer.WriteStartElement("wp", "posOffset", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteString(yEmu.ToString());
        _writer.WriteEndElement(); // wp:posOffset
        _writer.WriteEndElement(); // wp:positionV

        // Extent
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

        // Text wrapping
        if (anchor.WrapType == ShapeWrapType.None)
        {
            _writer.WriteStartElement("wp", "wrapNone", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            _writer.WriteEndElement();
        }
        else
        {
            _writer.WriteStartElement("wp", "wrapSquare", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            _writer.WriteAttributeString("wrapText", "bothSides");
            _writer.WriteEndElement();
        }

        // Doc properties
        _writer.WriteStartElement("wp", "docPr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteAttributeString("id", imageId.ToString());
        var baseName = !string.IsNullOrEmpty(image.FileName) ? image.FileName : $"Picture {imageId}";
        _writer.WriteAttributeString("name", baseName);
        var altText = baseName;
        var dotIndex = baseName.LastIndexOf('.');
        if (dotIndex > 0)
        {
            altText = baseName.Substring(0, dotIndex);
        }
        _writer.WriteAttributeString("descr", altText);
        _writer.WriteEndElement(); // wp:docPr

        // Non-visual graphic frame properties
        _writer.WriteStartElement("wp", "cNvGraphicFramePr", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteStartElement("a", "graphicFrameLocks", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("noChangeAspect", "1");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        // Graphic
        _writer.WriteStartElement("a", "graphic", "http://schemas.openxmlformats.org/drawingml/2006/main");
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
        _writer.WriteEndElement(); // pic:nvPicPr

        // Blip fill
        _writer.WriteStartElement("pic", "blipFill", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("a", "blip", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteAttributeString("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", relId);
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "stretch", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "fillRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // pic:blipFill

        // Shape properties
        _writer.WriteStartElement("pic", "spPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("a", "xfrm", "http://schemas.openxmlformats.org/drawingml/2006/main");
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
        _writer.WriteEndElement(); // pic:spPr

        _writer.WriteEndElement(); // pic:pic
        _writer.WriteEndElement(); // a:graphicData
        _writer.WriteEndElement(); // a:graphic
        _writer.WriteEndElement(); // wp:anchor
        _writer.WriteEndElement(); // w:drawing
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p
    }
    
    private void WriteSections(DocumentModel document)
    {
        // If there are explicit sections, their w:sectPr have already been written
        // inline after the corresponding paragraphs. For documents without any
        // SectionInfo, fall back to a single sectPr at the end of the body.
        if (document.Properties.Sections.Count == 0)
        {
            const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            _writer.WriteStartElement("w", "sectPr", wNs);
            WriteSectionContent(null);
            _writer.WriteEndElement();
        }
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

        // headerReference and footerReference must come first in sectPr
        var ids = RelationshipsWriter.ComputeRelationshipIds(_document);

        // Decide which logical header type (default/first/even/none) applies to this section.
        string? headerType = null;
        HeaderFooterReferenceType headerRef = section?.HeaderReference ?? HeaderFooterReferenceType.Default;

        headerType = headerRef switch
        {
            HeaderFooterReferenceType.Default => "default",
            HeaderFooterReferenceType.First => "first",
            HeaderFooterReferenceType.Even => "even",
            HeaderFooterReferenceType.None => null,
            _ => "default"
        };

        // Map logical header type to a concrete relationship ID.
        int headerRId = 0;
        if (headerType != null)
        {
            switch (headerRef)
            {
                case HeaderFooterReferenceType.First:
                    headerRId = ids.HeaderFirstRId != 0
                        ? ids.HeaderFirstRId
                        : (ids.HeaderOddRId != 0 ? ids.HeaderOddRId : ids.HeaderEvenRId);
                    break;
                case HeaderFooterReferenceType.Even:
                    headerRId = ids.HeaderEvenRId != 0
                        ? ids.HeaderEvenRId
                        : (ids.HeaderOddRId != 0 ? ids.HeaderOddRId : ids.HeaderFirstRId);
                    break;
                case HeaderFooterReferenceType.Default:
                default:
                    headerRId = ids.HeaderOddRId != 0
                        ? ids.HeaderOddRId
                        : (ids.HeaderFirstRId != 0 ? ids.HeaderFirstRId : ids.HeaderEvenRId);
                    break;
            }
        }

        if (headerType != null && headerRId > 0)
        {
            _writer.WriteStartElement("w", "headerReference", wNs);
            _writer.WriteAttributeString("w", "type", wNs, headerType);
            _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{headerRId}");
            _writer.WriteEndElement();
        }

        // Decide which logical footer type applies to this section.
        string? footerType = null;
        HeaderFooterReferenceType footerRef = section?.FooterReference ?? HeaderFooterReferenceType.Default;

        footerType = footerRef switch
        {
            HeaderFooterReferenceType.Default => "default",
            HeaderFooterReferenceType.First => "first",
            HeaderFooterReferenceType.Even => "even",
            HeaderFooterReferenceType.None => null,
            _ => "default"
        };

        int footerRId = 0;
        if (footerType != null)
        {
            switch (footerRef)
            {
                case HeaderFooterReferenceType.First:
                    footerRId = ids.FooterFirstRId != 0
                        ? ids.FooterFirstRId
                        : (ids.FooterOddRId != 0 ? ids.FooterOddRId : ids.FooterEvenRId);
                    break;
                case HeaderFooterReferenceType.Even:
                    footerRId = ids.FooterEvenRId != 0
                        ? ids.FooterEvenRId
                        : (ids.FooterOddRId != 0 ? ids.FooterOddRId : ids.FooterFirstRId);
                    break;
                case HeaderFooterReferenceType.Default:
                default:
                    footerRId = ids.FooterOddRId != 0
                        ? ids.FooterOddRId
                        : (ids.FooterFirstRId != 0 ? ids.FooterFirstRId : ids.FooterEvenRId);
                    break;
            }
        }

        if (footerType != null && footerRId > 0)
        {
            _writer.WriteStartElement("w", "footerReference", wNs);
            _writer.WriteAttributeString("w", "type", wNs, footerType);
            _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{footerRId}");
            _writer.WriteEndElement();
        }

        // Page size and margins: prefer per-section overrides when available
        int pageWidth;
        int pageHeight;
        bool isLandscape;
        int marginTop;
        int marginBottom;
        int marginLeft;
        int marginRight;

        if (section != null)
        {
            pageWidth = section.PageWidth > 0 ? section.PageWidth : props.PageWidth;
            pageHeight = section.PageHeight > 0 ? section.PageHeight : props.PageHeight;
            isLandscape = section.IsLandscape;
            marginTop = section.MarginTop != 0 ? section.MarginTop : props.MarginTop;
            marginBottom = section.MarginBottom != 0 ? section.MarginBottom : props.MarginBottom;
            marginLeft = section.MarginLeft != 0 ? section.MarginLeft : props.MarginLeft;
            marginRight = section.MarginRight != 0 ? section.MarginRight : props.MarginRight;
        }
        else
        {
            pageWidth = props.PageWidth;
            pageHeight = props.PageHeight;
            isLandscape = props.IsLandscape;
            marginTop = props.MarginTop;
            marginBottom = props.MarginBottom;
            marginLeft = props.MarginLeft;
            marginRight = props.MarginRight;
        }

        _writer.WriteStartElement("w", "pgSz", wNs);
        _writer.WriteAttributeString("w", "w", wNs, pageWidth.ToString());
        _writer.WriteAttributeString("w", "h", wNs, pageHeight.ToString());
        if (isLandscape)
        {
            _writer.WriteAttributeString("w", "orient", wNs, "landscape");
        }
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "pgMar", wNs);
        _writer.WriteAttributeString("w", "top", wNs, marginTop.ToString());
        _writer.WriteAttributeString("w", "right", wNs, marginRight.ToString());
        _writer.WriteAttributeString("w", "bottom", wNs, marginBottom.ToString());
        _writer.WriteAttributeString("w", "left", wNs, marginLeft.ToString());
        _writer.WriteAttributeString("w", "header", wNs, "720");
        _writer.WriteAttributeString("w", "footer", wNs, "720");
        _writer.WriteAttributeString("w", "gutter", wNs, "0");
        _writer.WriteEndElement();

        // Mirror margins (left/right swapped on facing pages) – driven by DOP flag.
        if (props.FMirrorMargins)
        {
            _writer.WriteStartElement("w", "mirrorMargins", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "1");
            _writer.WriteEndElement();
        }

        // Page numbering start (document-level only, for now)
        if (section == null && props.SectionStartPageNumber > 1)
        {
            _writer.WriteStartElement("w", "pgNumType", wNs);
            _writer.WriteAttributeString("w", "start", wNs, props.SectionStartPageNumber.ToString());
            _writer.WriteEndElement();
        }

        // Columns
        _writer.WriteStartElement("w", "cols", wNs);
        _writer.WriteAttributeString("w", "space", wNs, "720");
        _writer.WriteEndElement();
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

    /// <summary>
    /// Writes a table to the document.
    /// </summary>
    private void WriteTable(TableModel table)
    {
        _writer.WriteStartElement("w", "tbl", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Write table properties
        WriteTableProperties(table);
        
        _writer.WriteStartElement("w", "tblGrid", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        int columnCount = table.ColumnCount > 0
            ? table.ColumnCount
            : (table.Rows.Any() ? table.Rows.Max(r => r.Cells.Count) : 0);
        for (int i = 0; i < columnCount; i++)
        {
            _writer.WriteStartElement("w", "gridCol", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            int width = 0;
            if (table.Rows.Count > 0 && i < table.Rows[0].Cells.Count && table.Rows[0].Cells[i].Properties?.Width > 0)
            {
                width = table.Rows[0].Cells[i].Properties!.Width;
            }
            
            if (width > 0)
            {
                _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", width.ToString());
            }
            _writer.WriteEndElement();
        }
        _writer.WriteEndElement();

        // Write each row
        foreach (var row in table.Rows)
        {
            WriteTableRow(row, table);
        }
        
        _writer.WriteEndElement(); // w:tbl
    }

    /// <summary>
    /// Writes table properties (tblPr).
    /// </summary>
    private void WriteTableProperties(TableModel table)
    {
        _writer.WriteStartElement("w", "tblPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Table style
        if (table.Properties?.StyleIndex >= 0)
        {
            var style = _document?.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Table && s.StyleId == table.Properties.StyleIndex);
            var styleId = StyleHelper.GetTableStyleId(table.Properties.StyleIndex, style?.Name);
            
            _writer.WriteStartElement("w", "tblStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", styleId);
            _writer.WriteEndElement();
        }
        
        // Table width: prefer an explicit width from TAP when available, otherwise
        // let Word auto-size based on content.
        _writer.WriteStartElement("w", "tblW", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var preferredWidth = table.Properties?.PreferredWidth ?? 0;
        if (preferredWidth > 0)
        {
            _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", preferredWidth.ToString());
            _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        }
        else
        {
            _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0");
            _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "auto");
        }
        _writer.WriteEndElement();
        
        // Table justification (alignment)
        if (table.Properties != null && table.Properties.Alignment != TableAlignment.Left)
        {
            _writer.WriteStartElement("w", "jc", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var alignment = table.Properties.Alignment switch
            {
                TableAlignment.Center => "center",
                TableAlignment.Right => "right",
                _ => "left"
            };
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", alignment);
            _writer.WriteEndElement();
        }
        
        // Table indent from left margin, when specified. This mirrors sprmTDxaLeft
        // and helps nested or offset tables align closer to the original layout.
        if (table.Properties != null && table.Properties.Indent != 0)
        {
            _writer.WriteStartElement("w", "tblInd", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", table.Properties.Indent.ToString());
            _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
            _writer.WriteEndElement();
        }
        
        // Table borders
        if (table.Properties?.BorderTop != null || table.Properties?.BorderBottom != null ||
            table.Properties?.BorderLeft != null || table.Properties?.BorderRight != null ||
            table.Properties?.BorderInsideH != null || table.Properties?.BorderInsideV != null)
        {
            _writer.WriteStartElement("w", "tblBorders", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (table.Properties.BorderTop != null) WriteBorder("top", table.Properties.BorderTop);
            if (table.Properties.BorderBottom != null) WriteBorder("bottom", table.Properties.BorderBottom);
            if (table.Properties.BorderLeft != null) WriteBorder("left", table.Properties.BorderLeft);
            if (table.Properties.BorderRight != null) WriteBorder("right", table.Properties.BorderRight);
            if (table.Properties.BorderInsideH != null) WriteBorder("insideH", table.Properties.BorderInsideH);
            if (table.Properties.BorderInsideV != null) WriteBorder("insideV", table.Properties.BorderInsideV);
            _writer.WriteEndElement();
        }
        
        // Table shading
        if (table.Properties?.Shading != null)
        {
            WriteShading(table.Properties.Shading);
        }
        
        // Table cell margin: when the TAP exposes an inter-cell spacing we map it
        // to symmetric left/right padding; otherwise we fall back to a sensible
        // default that keeps existing documents visually similar.
        var spacing = table.Properties?.CellSpacing ?? 0;
        // Clamp to a small, non-negative range so extreme values from corrupted
        // documents do not explode layout.
        if (spacing < 0) spacing = 0;
        if (spacing > 720) spacing = 720; // max 0.5"

        int sidePadding = spacing > 0 ? spacing / 2 : 108;

        _writer.WriteStartElement("w", "tblCellMar", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "top", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0");
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "left", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", sidePadding.ToString());
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "bottom", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0");
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "right", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", sidePadding.ToString());
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement(); // w:tblPr
    }

    /// <summary>
    /// Writes a table row.
    /// </summary>
    private void WriteTableRow(TableRowModel row, TableModel table)
    {
        _writer.WriteStartElement("w", "tr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Row properties
        if (row.Properties != null)
        {
            _writer.WriteStartElement("w", "trPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            if (row.Properties.Height > 0)
            {
                _writer.WriteStartElement("w", "trHeight", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", row.Properties.Height.ToString());
                if (row.Properties.HeightIsExact)
                {
                    _writer.WriteAttributeString("w", "hRule", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "exact");
                }
                _writer.WriteEndElement();
            }
            
            if (row.Properties.IsHeaderRow)
            {
                _writer.WriteStartElement("w", "tblHeader", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteEndElement();
            }

            // Prevent row from being split across pages when requested
            if (!row.Properties.AllowBreakAcrossPages)
            {
                _writer.WriteStartElement("w", "cantSplit", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteEndElement();
            }
            
            _writer.WriteEndElement(); // w:trPr
        }
        
        // Write each cell
        foreach (var cell in row.Cells)
        {
            WriteTableCell(cell, row, table);
        }
        
        _writer.WriteEndElement(); // w:tr
    }

    /// <summary>
    /// Writes a table cell, including vertical (vMerge) and horizontal (gridSpan)
    /// merges. For vertical merges we emit w:vMerge restart/continue based on
    /// RowSpan and cells in previous rows; for horizontal merges we emit
    /// w:gridSpan on the first cell and suppress content in covered cells.
    /// </summary>
    private void WriteTableCell(TableCellModel cell, TableRowModel row, TableModel table)
    {
        _writer.WriteStartElement("w", "tc", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // Determine vertical merge role for this cell
        bool isVmergeStart = cell.RowSpan > 1;
        bool isVmergeContinue = !isVmergeStart && IsCoveredByVerticalMerge(table, row.Index, cell.ColumnIndex);
        bool isHmergeCovered = IsCoveredByHorizontalMerge(table, row.Index, cell.ColumnIndex);
        
        bool hasTcPr = cell.Properties?.Width > 0 || cell.ColumnSpan > 1 || cell.RowSpan > 1 || isVmergeContinue ||
                       cell.Properties?.BorderTop != null || cell.Properties?.BorderBottom != null ||
                       cell.Properties?.BorderLeft != null || cell.Properties?.BorderRight != null ||
                       cell.Properties?.NoWrap == true ||
                       (cell.Properties != null && cell.Properties.VerticalAlignment != VerticalAlignment.Top);

        if (hasTcPr)
        {
            // tcPr: tcW -> gridSpan -> vMerge -> tcBorders -> shd -> noWrap -> vAlign
            _writer.WriteStartElement("w", "tcPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            // Cell width
            if (cell.Properties?.Width > 0)
            {
                _writer.WriteStartElement("w", "tcW", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", cell.Properties.Width.ToString());
                _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
                _writer.WriteEndElement();
            }
            
            // Grid span (column span) — only on the first (uncovered) cell
            if (cell.ColumnSpan > 1 && !isHmergeCovered)
            {
                _writer.WriteStartElement("w", "gridSpan", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", cell.ColumnSpan.ToString());
                _writer.WriteEndElement();
            }
            
            // Vertical merge (row span)
            if (isVmergeStart || isVmergeContinue)
            {
                _writer.WriteStartElement("w", "vMerge", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                if (isVmergeStart)
                {
                    _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "restart");
                }
                _writer.WriteEndElement();
            }
            
            // Cell borders
            if (cell.Properties?.BorderTop != null || cell.Properties?.BorderBottom != null ||
                cell.Properties?.BorderLeft != null || cell.Properties?.BorderRight != null)
            {
                _writer.WriteStartElement("w", "tcBorders", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                if (cell.Properties.BorderTop != null) WriteBorder("top", cell.Properties.BorderTop);
                if (cell.Properties.BorderBottom != null) WriteBorder("bottom", cell.Properties.BorderBottom);
                if (cell.Properties.BorderLeft != null) WriteBorder("left", cell.Properties.BorderLeft);
                if (cell.Properties.BorderRight != null) WriteBorder("right", cell.Properties.BorderRight);
                _writer.WriteEndElement();
            }

            // Cell shading (shd)
            if (cell.Properties?.Shading != null)
            {
                WriteShading(cell.Properties.Shading);
            }

            // No wrap
            if (cell.Properties?.NoWrap == true)
            {
                _writer.WriteStartElement("w", "noWrap", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteEndElement();
            }

            // Vertical alignment
            if (cell.Properties != null && cell.Properties.VerticalAlignment != VerticalAlignment.Top)
            {
                _writer.WriteStartElement("w", "vAlign", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                var vAlign = cell.Properties.VerticalAlignment switch
                {
                    VerticalAlignment.Center => "center",
                    VerticalAlignment.Bottom => "bottom",
                    VerticalAlignment.Both => "both",
                    _ => "top"
                };
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", vAlign);
                _writer.WriteEndElement();
            }
            
            _writer.WriteEndElement(); // w:tcPr
        }
        
        // Write cell content (paragraphs) only for:
        //   - vertical-merge starting cells
        //   - horizontal-merge starting cells
        if (!IsCoveredByVerticalMerge(table, row.Index, cell.ColumnIndex) &&
            !isHmergeCovered)
        {
            foreach (var para in cell.Paragraphs)
            {
                WriteParagraph(para);
            }
        }
        
        _writer.WriteEndElement(); // w:tc
    }

    /// <summary>
    /// Returns true if the cell at (rowIndex, columnIndex) is within the vertical
    /// span of a cell above it (RowSpan &gt; 1).
    /// </summary>
    private static bool IsCoveredByVerticalMerge(TableModel table, int rowIndex, int columnIndex)
    {
        if (rowIndex <= 0) return false;

        for (int r = 0; r < rowIndex; r++)
        {
            var row = table.Rows[r];
            foreach (var c in row.Cells)
            {
                if (c.ColumnIndex != columnIndex) continue;
                if (c.RowSpan > 1)
                {
                    int start = c.RowIndex;
                    int end = c.RowIndex + c.RowSpan - 1;
                    if (rowIndex >= start && rowIndex <= end)
                    {
                        // This row is within the vertical span of the cell starting at 'start'
                        return rowIndex > start; // true for continuation rows only
                    }
                }
            }
        }

        return false;
    }
    
    /// <summary>
    /// Returns true if the cell at (rowIndex, columnIndex) is horizontally covered
    /// by a previous cell in the same row with ColumnSpan &gt; 1.
    /// </summary>
    private static bool IsCoveredByHorizontalMerge(TableModel table, int rowIndex, int columnIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count) return false;
        var row = table.Rows[rowIndex];
        if (columnIndex < 0 || columnIndex >= row.Cells.Count) return false;

        for (int c = 0; c < row.Cells.Count; c++)
        {
            var cell = row.Cells[c];
            if (cell.ColumnSpan > 1)
            {
                int spanStart = cell.ColumnIndex;
                int spanEnd = cell.ColumnIndex + cell.ColumnSpan - 1;
                if (columnIndex > spanStart && columnIndex <= spanEnd)
                {
                    return true;
                }
            }
        }

        return false;
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
            (!string.IsNullOrEmpty(r.Text) && !string.IsNullOrWhiteSpace(r.Text)) || r.IsPicture || r.IsField);
    }

    private void WriteParagraph(ParagraphModel paragraph, bool suppressPageBreakBefore = false)
    {
        // If this paragraph is actually a wrapper for a nested table, write the table directly
        if (paragraph.Type == ParagraphType.NestedTable && paragraph.NestedTable != null)
        {
            WriteTable(paragraph.NestedTable);
            return;
        }

        // Filter runs to only those with actual content
        var runsWithContent = paragraph.Runs.Where(r => !string.IsNullOrEmpty(r.Text) || r.IsPicture || r.IsField).ToList();
        
        // Always write the paragraph element - OOXML requires at least one w:p in table cells,
        // and empty paragraphs (blank lines, page breaks) are meaningful document structure.
        _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        if (paragraph.Properties != null)
        {
            WriteParagraphProperties(paragraph.Properties, suppressPageBreakBefore);
        }
        
        foreach (var run in runsWithContent)
        {
            WriteRun(run);
        }
        
        _writer.WriteEndElement(); // w:p
    }
    
    private void WriteParagraphProperties(ParagraphProperties props, bool suppressPageBreakBefore = false)
    {
        if (props == null) return;
        
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        
        // pPr sequence per ISO 29500 CT_PPr:
        // pStyle -> keepNext -> keepLines -> pageBreakBefore -> numPr -> pBdr -> shd ->
        // spacing -> ind -> jc -> outlineLvl
        _writer.WriteStartElement("w", "pPr", wNs);
        
        // 1. pStyle
        if (props.StyleIndex >= 0)
        {
            var style = _document?.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == props.StyleIndex);
            var styleId = StyleHelper.GetParagraphStyleId(props.StyleIndex, style?.Name);
            
            _writer.WriteStartElement("w", "pStyle", wNs);
            _writer.WriteAttributeString("w", "val", wNs, styleId);
            _writer.WriteEndElement();
        }

        // 2. keepNext
        if (props.KeepWithNext)
        {
            _writer.WriteStartElement("w", "keepNext", wNs);
            _writer.WriteEndElement();
        }
        
        // 3. keepLines
        if (props.KeepTogether)
        {
            _writer.WriteStartElement("w", "keepLines", wNs);
            _writer.WriteEndElement();
        }
        
        // 4. pageBreakBefore (suppressed at doc start so first content e.g. 绿色等级评价报告 stays on page 1)
        if (props.PageBreakBefore && !suppressPageBreakBefore)
        {
            _writer.WriteStartElement("w", "pageBreakBefore", wNs);
            _writer.WriteEndElement();
        }

        // 5. numPr
        if (props.ListFormatId > 0)
        {
            WriteNumberingProperties(props.ListFormatId, props.ListLevel);
        }

        // 6. pBdr
        if (props.BorderTop != null || props.BorderBottom != null || 
            props.BorderLeft != null || props.BorderRight != null)
        {
            _writer.WriteStartElement("w", "pBdr", wNs);
            if (props.BorderTop != null) WriteBorder("top", props.BorderTop);
            if (props.BorderBottom != null) WriteBorder("bottom", props.BorderBottom);
            if (props.BorderLeft != null) WriteBorder("left", props.BorderLeft);
            if (props.BorderRight != null) WriteBorder("right", props.BorderRight);
            _writer.WriteEndElement();
        }
        
        // 7. shd
        if (props.Shading != null)
        {
            WriteShading(props.Shading);
        }

        // 8. spacing
        if (props.SpaceBefore > 0 || props.SpaceAfter > 0 || props.LineSpacing > 0)
        {
            _writer.WriteStartElement("w", "spacing", wNs);
            if (props.SpaceBefore > 0)
                _writer.WriteAttributeString("w", "before", wNs, props.SpaceBefore.ToString());
            if (props.SpaceAfter > 0)
                _writer.WriteAttributeString("w", "after", wNs, props.SpaceAfter.ToString());
            if (props.LineSpacing > 0)
            {
                _writer.WriteAttributeString("w", "line", wNs, props.LineSpacing.ToString());
                if (props.LineSpacingMultiple > 0)
                    _writer.WriteAttributeString("w", "lineRule", wNs, props.LineSpacingMultiple == 1 ? "auto" : "exact");
                else
                    _writer.WriteAttributeString("w", "lineRule", wNs, "auto");
            }
            _writer.WriteEndElement();
        }
        
        // 9. ind
        if (props.IndentLeft != 0 || props.IndentRight != 0 || props.IndentFirstLine != 0)
        {
            _writer.WriteStartElement("w", "ind", wNs);
            if (props.IndentLeft != 0)
                _writer.WriteAttributeString("w", "left", wNs, props.IndentLeft.ToString());
            if (props.IndentRight != 0)
                _writer.WriteAttributeString("w", "right", wNs, props.IndentRight.ToString());
            
            if (props.IndentFirstLine > 0)
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
        if (props.Alignment != ParagraphAlignment.Left)
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
        if (props.OutlineLevel >= 0 && props.OutlineLevel < 9)
        {
            _writer.WriteStartElement("w", "outlineLvl", wNs);
            _writer.WriteAttributeString("w", "val", wNs, props.OutlineLevel.ToString());
            _writer.WriteEndElement();
        }

        // 12. Text Formatting / Typography Flags
        if (!props.WordWrap)
        {
            _writer.WriteStartElement("w", "wordWrap", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (!props.Kinsoku)
        {
            _writer.WriteStartElement("w", "kinsoku", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (!props.SnapToGrid)
        {
            _writer.WriteStartElement("w", "snapToGrid", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (!props.AutoSpaceDe)
        {
            _writer.WriteStartElement("w", "autoSpaceDE", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (!props.AutoSpaceDn)
        {
            _writer.WriteStartElement("w", "autoSpaceDN", wNs);
            _writer.WriteAttributeString("w", "val", wNs, "0");
            _writer.WriteEndElement();
        }
        if (props.TopLinePunct)
        {
            _writer.WriteStartElement("w", "topLinePunct", wNs);
            _writer.WriteEndElement();
        }
        if (props.OverflowPunct)
        {
            _writer.WriteStartElement("w", "overflowPunct", wNs);
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
        if (border.Style == BorderStyle.None) return;
        
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", position, wNs);
        _writer.WriteAttributeString("w", "val", wNs, GetBorderStyle(border.Style));
        _writer.WriteAttributeString("w", "sz", wNs, (border.Width / 8).ToString());
        _writer.WriteAttributeString("w", "space", wNs, "0");
        _writer.WriteAttributeString("w", "color", wNs, ColorHelper.ColorToHex(border.Color));
        _writer.WriteEndElement();
    }
    
    private void WriteShading(ShadingInfo shading)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "shd", wNs);
        // Use PatternVal (from SHD ipat) when set; otherwise map Pattern enum to OOXML val so pattern/tiled background is preserved
        var val = !string.IsNullOrEmpty(shading.PatternVal)
            ? shading.PatternVal
            : ShadingPatternToShdVal(shading.Pattern);
        _writer.WriteAttributeString("w", "val", wNs, val);
        if (shading.ForegroundColor != 0)
            _writer.WriteAttributeString("w", "color", wNs, ColorHelper.ColorToHex(shading.ForegroundColor));
        _writer.WriteAttributeString("w", "fill", wNs, ColorHelper.ColorToHex(shading.BackgroundColor));
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
        _writer.WriteStartElement("w", type, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", (_trackChangeId++).ToString());
        
        string author = "Unknown Author";
        if (type == "ins" && !string.IsNullOrEmpty(props.AuthorIndexIns.ToString())) author = props.AuthorIndexIns.ToString();
        else if (type == "del" && !string.IsNullOrEmpty(props.AuthorIndexDel.ToString())) author = props.AuthorIndexDel.ToString();
        _writer.WriteAttributeString("w", "author", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", author);
        
        uint dttm = type == "ins" ? props.DateIns : props.DateDel;
        if (dttm != 0)
        {
            try {
                int mint = (int)(dttm & 0x3F);
                int hr = (int)((dttm >> 6) & 0x1F);
                int dom = (int)((dttm >> 11) & 0x1F);
                int mon = (int)((dttm >> 16) & 0x0F);
                int yr = 1900 + (int)((dttm >> 20) & 0x1FF);
                var dt = new DateTime(yr, Math.Max(1, mon), Math.Max(1, dom), hr, mint, 0);
                _writer.WriteAttributeString("w", "date", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", dt.ToString("yyyy-MM-ddTHH:mm:ssZ"));
            } catch { }
        }
    }

    private void WriteRun(RunModel run)
    {
        // Skip runs with no content at all (no text, no picture, no field)
        bool hasText = !string.IsNullOrEmpty(run.Text);
        bool hasVisualContent = run.IsPicture || run.IsField;
        
        if (!hasText && !hasVisualContent)
        {
            // Even if no text, if there are properties, we might want to write them
            // But for now, skip empty runs to avoid corruption
            return;
        }
        
        // Handle bookmark start
        if (run.IsBookmark && run.IsBookmarkStart)
        {
            WriteBookmarkStart(run.BookmarkName);
        }

        // Handle hyperlink
        if (run.IsHyperlink && !string.IsNullOrEmpty(run.HyperlinkUrl))
        {
            WriteHyperlink(run);
        }
        else
        {
            _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            WriteRunProperties(run);

            if (run.IsPicture && run.ImageIndex >= 0)
            {
                WritePicture(run);
                _writer.WriteEndElement(); // w:r
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
                    _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteStartElement("w", "instrText", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                    _writer.WriteString(run.FieldCode);
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
    }

    /// <summary>
    /// Writes a hyperlink element (w:hyperlink).
    /// </summary>
    private void WriteHyperlink(RunModel run)
    {
        _writer.WriteStartElement("w", "hyperlink", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        if (!string.IsNullOrEmpty(run.HyperlinkRelationshipId))
        {
            _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", run.HyperlinkRelationshipId);
        }
        else if (!string.IsNullOrEmpty(run.HyperlinkUrl))
        {
            // For internal bookmarks
            if (run.HyperlinkUrl.StartsWith("#"))
            {
                _writer.WriteAttributeString("w", "anchor", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", run.HyperlinkUrl.Substring(1));
            }
            else
            {
                _writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", run.HyperlinkRelationshipId ?? "rIdHyperlink");
            }
        }

        // Only write run if there's text content
        if (!string.IsNullOrEmpty(run.Text))
        {
            _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            WriteRunProperties(run);
            WriteRunText(run);
            _writer.WriteEndElement(); // w:r
        }
        
        _writer.WriteEndElement(); // w:hyperlink
    }

    /// <summary>
    /// Writes a bookmark start element.
    /// </summary>
    private void WriteBookmarkStart(string? name)
    {
        if (string.IsNullOrEmpty(name)) return;

        if (!_bookmarkIds.TryGetValue(name, out var id))
        {
            id = ++_bookmarkCounter;
            _bookmarkIds[name] = id;
        }

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

        if (!_bookmarkIds.TryGetValue(name, out var id))
        {
            // If we never saw a start for this bookmark name, allocate one so
            // the resulting document remains structurally valid.
            id = ++_bookmarkCounter;
            _bookmarkIds[name] = id;
        }

        _writer.WriteStartElement("w", "bookmarkEnd", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", id.ToString());
        _writer.WriteEndElement();
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
        
        // Calculate relationship ID using shared logic
        var ids = RelationshipsWriter.ComputeRelationshipIds(_document);
        var relId = $"rId{ids.FirstImageRId + run.ImageIndex}";
        
        // Use actual image dimensions or sensible defaults
        var widthEmu = image.WidthEMU > 0 ? image.WidthEMU : 5715000; // Default ~6 inches
        var heightEmu = image.HeightEMU > 0 ? image.HeightEMU : 3810000; // Default ~4 inches

        // Respect per-image scale factors when present (100000 = 100%)
        if (image.ScaleX > 0 && image.ScaleX != 100000)
        {
            widthEmu = (int)(widthEmu * (image.ScaleX / 100000.0));
        }
        if (image.ScaleY > 0 && image.ScaleY != 100000)
        {
            heightEmu = (int)(heightEmu * (image.ScaleY / 100000.0));
        }

        // Full-page background: first picture in body always gets full page; else if size ≈ page use full page; else clamp to content
        const int emuPerTwip = 635; // 1 twip = 1/1440 inch; 1 inch = 914400 EMUs
        if (_document?.Properties != null)
        {
            var page = _document.Properties;
            int pageWidthEmu = page.PageWidth * emuPerTwip;
            int pageHeightEmu = page.PageHeight * emuPerTwip;
            bool forceFirstFullPage = _firstBodyPictureNotYetWritten && pageWidthEmu > 0 && pageHeightEmu > 0;
            bool looksFullPage = !forceFirstFullPage && (pageWidthEmu > 0 && pageHeightEmu > 0) &&
                (widthEmu >= pageWidthEmu * 0.85 || heightEmu >= pageHeightEmu * 0.85);
            if (forceFirstFullPage || looksFullPage)
            {
                widthEmu = pageWidthEmu;
                heightEmu = pageHeightEmu;
                if (forceFirstFullPage) _firstBodyPictureNotYetWritten = false;
            }
            else
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
        _writer.WriteAttributeString("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", relId);
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "stretch", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteStartElement("a", "fillRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // pic:blipFill
        
        // Shape properties
        _writer.WriteStartElement("pic", "spPr", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        _writer.WriteStartElement("a", "xfrm", "http://schemas.openxmlformats.org/drawingml/2006/main");
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
        _writer.WriteAttributeString("behindDoc", "0");
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

        // Wrap None (floating) or Wrap Square
        _writer.WriteStartElement("wp", "wrapNone", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        _writer.WriteEndElement();

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
    
    private void WriteRunProperties(RunModel run)
    {
        var props = run.Properties;
        if (props == null) return;
        
        if (!HasRunProperties(props)) return;
        
        // rPr sequence: rStyle -> rFonts -> b -> bCs -> i -> iCs -> caps -> smallCaps -> strike -> outline -> shadow -> emboss -> color -> kern -> sz -> szCs -> highlight -> u -> vertAlign -> lang
        _writer.WriteStartElement("w", "rPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // 1. rFonts
        if (!string.IsNullOrEmpty(props.FontName))
        {
            _writer.WriteStartElement("w", "rFonts", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "ascii", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.FontName);
            _writer.WriteAttributeString("w", "hAnsi", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.FontName);
            _writer.WriteAttributeString("w", "cs", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.FontName);
            _writer.WriteEndElement();
        }
        
        // 2. b / bCs
        if (props.IsBold)
        {
            _writer.WriteStartElement("w", "b", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsBoldCs)
        {
            _writer.WriteStartElement("w", "bCs", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        
        // 3. i / iCs
        if (props.IsItalic)
        {
            _writer.WriteStartElement("w", "i", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsItalicCs)
        {
            _writer.WriteStartElement("w", "iCs", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        
        // 4. caps / smallCaps
        if (props.IsAllCaps)
        {
            _writer.WriteStartElement("w", "caps", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsSmallCaps)
        {
            _writer.WriteStartElement("w", "smallCaps", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }

        // 5. strike
        if (props.IsStrikeThrough || props.IsDoubleStrikeThrough)
        {
            _writer.WriteStartElement("w", "strike", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }

        // 5.5 hidden text
        if (props.IsHidden)
        {
            _writer.WriteStartElement("w", "vanish", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }

        // 6. outline / shadow / emboss / imprint
        if (props.IsOutline)
        {
            _writer.WriteStartElement("w", "outline", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsShadow)
        {
            _writer.WriteStartElement("w", "shadow", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsEmboss)
        {
            _writer.WriteStartElement("w", "emboss", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        if (props.IsImprint)
        {
            _writer.WriteStartElement("w", "imprint", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }

        // 7. color
        if (props.Color != 0 || props.HasRgbColor)
        {
            string colorHex;
            if (props.HasRgbColor)
            {
                colorHex = ColorHelper.RgbToHex(props.RgbColor);
            }
            else
            {
                colorHex = ColorHelper.ColorToHex(props.Color);
            }
            if (colorHex != "auto")
            {
                _writer.WriteStartElement("w", "color", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", colorHex);
                _writer.WriteEndElement();
            }
        }

        // 8. kern
        if (props.Kerning > 0)
        {
             _writer.WriteStartElement("w", "kern", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
             _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.Kerning.ToString());
             _writer.WriteEndElement();
        }

        // 9. spacing (character spacing)
        if (props.CharacterSpacingAdjustment != 0)
        {
            _writer.WriteStartElement("w", "spacing", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.CharacterSpacingAdjustment.ToString());
            _writer.WriteEndElement();
        }

        // 10. sz / szCs
        if (props.FontSize > 0 && props.FontSize != 24)
        {
            _writer.WriteStartElement("w", "sz", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.FontSize.ToString());
            _writer.WriteEndElement();
        }
        if (props.FontSizeCs > 0 && props.FontSizeCs != 24)
        {
            _writer.WriteStartElement("w", "szCs", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.FontSizeCs.ToString());
            _writer.WriteEndElement();
        }
        
        // 11. highlight
        if (props.HighlightColor > 0)
        {
            _writer.WriteStartElement("w", "highlight", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", ColorHelper.GetHighlightName(props.HighlightColor));
            _writer.WriteEndElement();
        }
        
        // 12. u
        if (props.IsUnderline)
        {
            _writer.WriteStartElement("w", "u", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", GetUnderlineType(props.UnderlineType));
            _writer.WriteEndElement();
        }
        
        // 13. vertAlign / position
        if (props.IsSuperscript)
        {
            _writer.WriteStartElement("w", "vertAlign", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "superscript");
            _writer.WriteEndElement();
        }
        else if (props.IsSubscript)
        {
            _writer.WriteStartElement("w", "vertAlign", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "subscript");
            _writer.WriteEndElement();
        }

        // Explicit position offset (in half-points)
        if (props.Position != 0 && !props.IsSuperscript && !props.IsSubscript)
        {
            _writer.WriteStartElement("w", "position", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.Position.ToString());
            _writer.WriteEndElement();
        }

        // 13.5 Character Scale (w)
        if (props.CharacterScale != 100 && props.CharacterScale > 0)
        {
            _writer.WriteStartElement("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.CharacterScale.ToString());
            _writer.WriteEndElement();
        }

        // 13.6 Snap to Grid
        if (!props.SnapToGrid)
        {
            _writer.WriteStartElement("w", "snapToGrid", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0");
            _writer.WriteEndElement();
        }

        // 13. lang
        if (props.Language > 0 || !string.IsNullOrEmpty(props.LanguageAsia) || !string.IsNullOrEmpty(props.LanguageCs))
        {
            _writer.WriteStartElement("w", "lang", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (props.Language > 0)
            {
                // Language is stored as LCID; map common values to BCP-47 where possible, otherwise omit.
                var lang = props.Language switch
                {
                    0x0409 => "en-US",
                    0x0804 => "zh-CN",
                    0x0404 => "zh-TW",
                    0x0411 => "ja-JP",
                    0x0412 => "ko-KR",
                    0x0407 => "de-DE",
                    0x040C => "fr-FR",
                    0x0410 => "it-IT",
                    0x0C0A => "es-ES",
                    _ => null
                };
                if (lang != null)
                {
                    _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", lang);
                }
            }

            if (!string.IsNullOrEmpty(props.LanguageAsia))
            {
                _writer.WriteAttributeString("w", "eastAsia", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.LanguageAsia);
            }
            if (!string.IsNullOrEmpty(props.LanguageCs))
            {
                _writer.WriteAttributeString("w", "bidi", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", props.LanguageCs);
            }
            _writer.WriteEndElement();
        }

        _writer.WriteEndElement(); // w:rPr
    }
    
    private bool HasRunProperties(RunProperties props)
    {
        return props.IsBold || props.IsBoldCs || props.IsItalic || props.IsItalicCs ||
               props.IsUnderline || props.IsStrikeThrough || props.IsDoubleStrikeThrough ||
               props.IsSmallCaps || props.IsAllCaps || props.IsSuperscript || props.IsSubscript ||
               props.FontSize != 24 || props.Color != 0 || !string.IsNullOrEmpty(props.FontName);
    }
    
    private string GetUnderlineType(UnderlineType type)
    {
        return type switch
        {
            UnderlineType.Single => "single",
            UnderlineType.WordsOnly => "word",
            UnderlineType.Double => "double",
            UnderlineType.Dotted => "dotted",
            UnderlineType.Thick => "thick",
            UnderlineType.Dash => "dash",
            UnderlineType.DotDash => "dotDash",
            UnderlineType.DotDotDash => "dotDotDash",
            UnderlineType.Wave => "wave",
            UnderlineType.ThickWave => "thickWave",
            _ => "none"
        };
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
        string text = run.Text.Replace("\r\n", "\n").Replace("\r", "\n");

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
                            _writer.WriteAttributeString("xml", "space", null, "preserve");
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
    
    /// <summary>
    /// Removes characters that are invalid in XML 1.0 documents and replaces
    /// U+FFFD (replacement character) with space to avoid black squares in Word.
    /// Valid: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
    /// </summary>
    private static string SanitizeXmlString(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new System.Text.StringBuilder(text.Length);
        foreach (char c in text)
        {
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
            // else: skip invalid XML character
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

