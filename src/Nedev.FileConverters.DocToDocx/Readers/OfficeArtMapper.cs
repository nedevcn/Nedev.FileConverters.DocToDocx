using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Maps low-level Escher (OfficeArt) records into high-level ShapeModel instances.
/// 当前阶段仅做基础形状发现和 Id 提取，后续阶段再补充锚点、图片和样式信息。
/// </summary>
public static class OfficeArtMapper
{
    // Escher record type constants (subset)
    private const ushort RecordTypeSpContainer = 0xF004;
    private const ushort RecordTypeSp = 0xF00A;

    public static void AttachShapes(DocumentModel document, OfficeArtReader? officeArtReader, IReadOnlyList<FspaInfo>? fspaAnchors)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (officeArtReader == null) return;
        if (officeArtReader.RootRecords.Count == 0) return;

        var shapes = new List<ShapeModel>();
        var imageIndexCursor = 0;

        foreach (var root in officeArtReader.RootRecords)
        {
            Traverse(root, shapes, document, ref imageIndexCursor);
        }

        // 从 FSPA 中为形状附加锚点信息（位置/大小），并结合 CP 映射到段落。
        if (fspaAnchors != null && fspaAnchors.Count > 0)
        {
            AttachAnchorsFromFspa(document, shapes, fspaAnchors);
        }
        else
        {
            // 如果没有 FSPA 信息，则回退到基于段落分布的启发式。
            AssignParagraphHints(document, shapes);
        }

        if (shapes.Count > 0)
        {
            document.Shapes.AddRange(shapes);
        }
    }

    private static void Traverse(EscherRecord record, List<ShapeModel> shapes, DocumentModel document, ref int imageIndexCursor)
    {
        if (record.Type == RecordTypeSpContainer)
        {
            var shape = CreateShapeFromSpContainer(record, document, ref imageIndexCursor);
            if (shape != null)
            {
                shapes.Add(shape);
            }
        }

        if (record.Children.Count == 0) return;
        foreach (var child in record.Children)
        {
            Traverse(child, shapes, document, ref imageIndexCursor);
        }
    }

    private static void AssignParagraphHints(DocumentModel document, List<ShapeModel> shapes)
    {
        if (shapes.Count == 0 || document.Paragraphs.Count == 0) return;

        var pictureShapes = shapes
            .Where(s => s.Type == ShapeType.Picture && s.ImageIndex is not null)
            .ToList();
        if (pictureShapes.Count == 0) return;

        // 优先选择普通段落作为锚点候选；如果没有，则退回全部段落。
        var candidateParagraphIndices = document.Paragraphs
            .Where(p => p.Type == ParagraphType.Normal)
            .Select(p => p.Index)
            .ToList();

        if (candidateParagraphIndices.Count == 0)
        {
            candidateParagraphIndices = document.Paragraphs.Select(p => p.Index).ToList();
        }

        if (candidateParagraphIndices.Count == 0) return;

        // 将图片形状均匀分布到候选段落索引上，作为“段落位置提示”。
        for (int i = 0; i < pictureShapes.Count; i++)
        {
            var target = candidateParagraphIndices[(int)((long)i * candidateParagraphIndices.Count / pictureShapes.Count)];
            pictureShapes[i].ParagraphIndexHint = target;
        }
    }

    /// <summary>
    /// Best-effort extraction of a shape from an SpContainer: we look for the
    /// EscherSp record, use its shapeId as the ShapeModel.Id, and map the
    /// MSOSPT shape type (from the Escher header instance) into a coarse
    /// ShapeType (rectangle, ellipse, textbox, picture). When we detect a
    /// picture frame we opportunistically associate it with the next extracted
    /// ImageModel so that OfficeArt picture shapes line up with decoded BLIPs.
    /// </summary>
    private static ShapeModel? CreateShapeFromSpContainer(EscherRecord spContainer, DocumentModel document, ref int imageIndexCursor)
    {
        ShapeModel? shape = null;
        int? shapeId = null;
        ShapeType mappedType = ShapeType.Unknown;

        foreach (var child in spContainer.Children)
        {
            if (child.Type == RecordTypeSp && child.Data.Length >= 4)
            {
                shapeId = BitConverter.ToInt32(child.Data, 0);
                mappedType = MapMsosptToShapeType(child.Instance);
            }
            else if (child.Type == 0xF00B) // RecordTypeOpt
            {
                // We'll parse properties from the OPT record later if we have a shape.
            }
        }

        if (shapeId == null) return null;

        int? imageIndex = null;
        if (mappedType == ShapeType.Picture && imageIndexCursor < document.Images.Count)
        {
            imageIndex = imageIndexCursor++;
        }
        else if (mappedType == ShapeType.Unknown && imageIndexCursor < document.Images.Count)
        {
            mappedType = ShapeType.Picture;
            imageIndex = imageIndexCursor++;
        }

        shape = new ShapeModel
        {
            Id = shapeId.Value,
            Type = mappedType,
            ImageIndex = imageIndex,
            IsLineVisible = true
        };

        // Parse OPT properties if present
        var opt = spContainer.Children.FirstOrDefault(c => c.Type == 0xF00B);
        if (opt != null)
        {
            ParseOptProperties(opt.Data, shape, spContainer);
        }

        return shape;
    }

    private static void ParseOptProperties(byte[] data, ShapeModel shape, EscherRecord? container = null)
    {
        // MS-ODRAW 2.1.1: OfficeArtOPT record
        // Each property is 6 bytes: 2 bytes ID (with flags) + 4 bytes value
        int pos = 0;
        int count = data.Length / 6;

        for (int i = 0; i < count; i++)
        {
            if (pos + 6 > data.Length) break;
            ushort propIdWithFlags = BitConverter.ToUInt16(data, pos);
            uint propValue = BitConverter.ToUInt32(data, pos + 2);
            pos += 6;

            ushort propId = (ushort)(propIdWithFlags & 0x3FFF);
            bool fComplex = (propIdWithFlags & 0x8000) != 0;

            switch (propId)
            {
                // Cropping (16.16 fixed point)
                case 257: shape.CropTop = (int)propValue; break;
                case 258: shape.CropBottom = (int)propValue; break;
                case 259: shape.CropLeft = (int)propValue; break;
                case 260: shape.CropRight = (int)propValue; break;

                // Line properties
                case 448: shape.LineColor = (int)propValue; break;
                case 459: shape.LineWidth = (int)propValue; break;
                case 511: shape.IsLineVisible = (propValue & 0x01) != 0; break;

                // Fill properties
                case 384: shape.FillColor = (int)propValue; break;

                // Custom Geometry properties [Phase 4.3]
                case 321: // pVertices (complex)
                    if (fComplex)
                    {
                        var vertexData = ParseComplexProperty(container, i, propId);
                        if (vertexData != null) ParseVertices(vertexData, shape);
                    }
                    break;
                case 322: // pSegmentInfo (complex)
                    if (fComplex)
                    {
                        var segmentData = ParseComplexProperty(container, i, propId);
                        if (segmentData != null) ParseSegments(segmentData, shape);
                    }
                    break;
                case 324: // pGeoTop
                    EnsureCustomGeometry(shape).ViewTop = (int)propValue;
                    break;
                case 323: // pGeoLeft
                    EnsureCustomGeometry(shape).ViewLeft = (int)propValue;
                    break;
                case 325: // pGeoRight
                    EnsureCustomGeometry(shape).ViewRight = (int)propValue;
                    break;
                case 326: // pGeoBottom
                    EnsureCustomGeometry(shape).ViewBottom = (int)propValue;
                    break;
            }
        }
    }

    private static CustomGeometry EnsureCustomGeometry(ShapeModel shape)
    {
        if (shape.CustomGeometry == null)
        {
            shape.CustomGeometry = new CustomGeometry();
            shape.Type = ShapeType.Custom;
        }
        return shape.CustomGeometry;
    }

    private static byte[]? ParseComplexProperty(EscherRecord? container, int propertyIndex, ushort propId)
    {
        if (container == null) return null;
        // Complex property data is stored in the 0xF011 (RecordTypeTertiaryOpt) or 
        // 0xF00B (RecordTypeOpt) record's tail or in a subsequent record.
        // Actually, Escher standard: if fComplex is set, the propValue IS the size (count), 
        // and the data follows the array of property headers.
        
        // This is a simplification: in simple Opt records, the complex data starts after the 
        // fixed-size (6-byte) property headers.
        // We'll look at the container's Opt record data.
        var opt = container.Children.FirstOrDefault(c => c.Type == 0xF00B || c.Type == 0xF011);
        if (opt == null) return null;

        // Find how many properties are defined to skip the headers
        int propCount = BitConverter.ToUInt16(opt.Data, 0); // Oops, Opt doesn't have a 2-byte count at the start!
        // Wait, MS-ODRAW 2.1.1: The OfficeArtOPT record payload is an array of OfficeArtOPTEntry.
        // The first OfficeArtOPTEntry.opid with bit 24 set to 1 indicates the start of complex data? No.
        
        // Actually, the simplest way is to look at the 'instance' of the Opt record - it's the count.
        int totalProps = opt.Instance;
        int headersSize = totalProps * 6;
        
        // Complex data starts after all headers. But which complex data?
        // They are stored in order of their appearance in the headers with fComplex set.
        int complexOffset = headersSize;
        int complexDataIndex = 0;
        
        for (int i = 0; i < totalProps; i++)
        {
            ushort entryIdFlags = BitConverter.ToUInt16(opt.Data, i * 6);
            uint entryValue = BitConverter.ToUInt32(opt.Data, i * 6 + 2);
            ushort entryId = (ushort)(entryIdFlags & 0x3FFF);
            bool entryComplex = (entryIdFlags & 0x8000) != 0;

            if (entryComplex)
            {
                if (entryId == propId) // This is the one we want
                {
                    if (complexOffset + entryValue > opt.Data.Length) return null;
                    byte[] complexData = new byte[entryValue];
                    Array.Copy(opt.Data, complexOffset, complexData, 0, (int)entryValue);
                    return complexData;
                }
                complexOffset += (int)entryValue;
            }
        }

        return null;
    }

    private static void ParseVertices(byte[] data, ShapeModel shape)
    {
        if (data.Length < 6) return; // Header (sizeof(struct array))
        
        int nElems = BitConverter.ToUInt16(data, 0);
        int nElemsAlloc = BitConverter.ToUInt16(data, 2);
        int cbElem = BitConverter.ToUInt16(data, 4); // should be 8 for Point

        var geom = EnsureCustomGeometry(shape);
        int pos = 6;
        for (int i = 0; i < nElems && pos + 8 <= data.Length; i++)
        {
            int x = BitConverter.ToInt32(data, pos);
            int y = BitConverter.ToInt32(data, pos + 4);
            geom.Vertices.Add(new System.Drawing.Point(x, y));
            pos += 8;
        }
    }

    private static void ParseSegments(byte[] data, ShapeModel shape)
    {
        if (data.Length < 6) return;
        
        int nElems = BitConverter.ToUInt16(data, 0);
        int cbElem = BitConverter.ToUInt16(data, 4); // usually 2 for segments

        var geom = EnsureCustomGeometry(shape);
        int pos = 6;
        int vertexIdx = 0;
        for (int i = 0; i < nElems && pos + 2 <= data.Length; i++)
        {
            ushort code = BitConverter.ToUInt16(data, pos);
            pos += 2;

            // MS-ODRAW 2.1.18: Table 33 - Shape Path Segment Codes
            // Simplified mapping:
            if (code >= 0x0000 && code <= 0x00A7) 
            {
                // This is a lineTo or moveTo depending on context
                // But specifically for Word, segment info codes usually:
                // 0x0001: lineto, 0x0003: moveto, 0x00BE: curveTo, 0x8000: close
                
                // Let's use a more robust logic if we find better documentation or samples.
                // For now:
                var segment = new ShapePathSegment { VertexIndex = vertexIdx };
                if (code == 0x0001) { segment.Type = SegmentType.LineTo; vertexIdx += 1; }
                else if (code == 0x0002) { segment.Type = SegmentType.CurveTo; vertexIdx += 3; }
                else if (code == 0x0003) { segment.Type = SegmentType.MoveTo; vertexIdx += 1; }
                else if (code == 0x2001) { segment.Type = SegmentType.Close; }
                else if (code == 0x00AA) { segment.Type = SegmentType.End; }
                else { segment.Type = SegmentType.LineTo; vertexIdx += 1; } // Fallback
                
                geom.Segments.Add(segment);
            }
        }
    }

    /// <summary>
    /// Maps MSOSPT (preset shape type from MS-ODRAW, see MSOSPT enum) to our
    /// coarse ShapeType. This is intentionally lossy but helps distinguish
    /// rectangles, ellipses, textboxes and picture frames for better visual
    /// fidelity and SmartArt fallbacks.
    /// </summary>
    private static ShapeType MapMsosptToShapeType(ushort msospt)
    {
        return msospt switch
        {
            // Basic rectangles and rectangle-like shapes
            0x0001 or // msosptRectangle
            0x0002 or // msosptRoundRectangle
            0x0004 or // msosptDiamond
            0x0007 or // msosptParallelogram
            0x0008 or // msosptTrapezoid
            0x0015 or // msosptPlaque
            0x0041 or // msosptFoldedCorner
            0x0054 or // msosptBevel
            0x006D or // msosptFlowChartProcess
            0x0072 or // msosptFlowChartDocument
            0x0074 or // msosptFlowChartTerminator
            0x00B0 or // msosptFlowChartAlternateProcess
            0x00B1     // msosptFlowChartOffpageConnector
                => ShapeType.Rectangle,

            // Ellipses and ellipse-like
            0x0003 or // msosptEllipse
            0x0003F   // msosptWedgeEllipseCallout
                => ShapeType.Ellipse,

            // Text boxes
            0x00CA // msosptTextBox
                => ShapeType.Textbox,

            // Picture frame
            0x004B // msosptPictureFrame
                => ShapeType.Picture,

            _ => ShapeType.Unknown
        };
    }

    /// <summary>
    /// Uses FSPA anchors to populate ShapeAnchor (floating position/size) and
    /// ParagraphIndexHint based on CP values.
    /// </summary>
    private static void AttachAnchorsFromFspa(DocumentModel document, List<ShapeModel> shapes, IReadOnlyList<FspaInfo> fspaAnchors)
    {
        if (shapes.Count == 0 || fspaAnchors.Count == 0 || document.Paragraphs.Count == 0)
            return;

        // Build a quick lookup from spid to FSPA info (last one wins if duplicates).
        var fspaBySpid = new Dictionary<int, FspaInfo>();
        foreach (var fspa in fspaAnchors)
        {
            fspaBySpid[fspa.Spid] = fspa;
        }

        // Precompute paragraphs sorted by minimum CP (CharacterPosition) to
        // approximate where shapes should be attached.
        var paraInfos = document.Paragraphs
            .Select(p => new
            {
                Paragraph = p,
                MinCp = p.Runs.Count > 0 ? p.Runs.Min(r => r.CharacterPosition) : int.MaxValue
            })
            .OrderBy(p => p.MinCp)
            .ToList();

        int zOrderCounter = 0;

        foreach (var shape in shapes)
        {
            if (!fspaBySpid.TryGetValue(shape.Id, out var fspa))
                continue;

            // Populate anchor position and size from the FSPA bounding box.
            var width = fspa.XaRight - fspa.XaLeft;
            var height = fspa.YaBottom - fspa.YaTop;
            if (width <= 0 || height <= 0)
                continue;

            shape.Anchor = new ShapeAnchor
            {
                IsFloating = true,
                PageIndex = 0,
                ParagraphIndex = -1,
                X = fspa.XaLeft,
                Y = fspa.YaTop,
                Width = width,
                Height = height,
                HorizontalRelativeTo = MapRelativeToHorizontal(fspa.Flags),
                VerticalRelativeTo = MapRelativeToVertical(fspa.Flags),
                WrapType = MapWrapType(fspa.Flags),
                ZOrder = zOrderCounter++
            };

            // Map CP to the paragraph containing the CP.
            // A paragraph contains the CP if its MinCp is <= the shape's CP.
            // Since paraInfos is ordered by MinCp ascending, the LAST one satisfying MinCp <= cp is the container.
            var cp = fspa.Cp;
            var bestPara = paraInfos.LastOrDefault(p => p.MinCp != int.MaxValue && p.MinCp <= cp);
            if (bestPara == null)
            {
                bestPara = paraInfos.FirstOrDefault(p => p.MinCp != int.MaxValue);
            }

            if (bestPara != null)
            {
                shape.ParagraphIndexHint = bestPara.Paragraph.Index;
                shape.Anchor.ParagraphIndex = bestPara.Paragraph.Index;
            }
        }
    }

    /// <summary>
    /// Maps FSPA flags to a horizontal reference frame. This is intentionally
    /// conservative: until all flag combinations are well understood, we default
    /// to page-relative anchors and only special-case a few common patterns.
    /// </summary>
    private static ShapeRelativeTo MapRelativeToHorizontal(ushort flags)
    {
        // bits 1-2: bx (0=page, 1=margin, 2=column, 3=char)
        int bx = (flags >> 1) & 0x03;
        return bx switch
        {
            1 => ShapeRelativeTo.Margin,
            2 => ShapeRelativeTo.Column,
            3 => ShapeRelativeTo.Column, // character not fully supported, map to column
            _ => ShapeRelativeTo.Page
        };
    }

    /// <summary>
    /// Maps FSPA flags to a vertical reference frame. See comments on
    /// MapRelativeToHorizontal for caveats.
    /// </summary>
    private static ShapeRelativeTo MapRelativeToVertical(ushort flags)
    {
        // bits 3-4: by (0=page, 1=margin, 2=paragraph, 3=line)
        int by = (flags >> 3) & 0x03;
        return by switch
        {
            1 => ShapeRelativeTo.Margin,
            2 => ShapeRelativeTo.Paragraph,
            3 => ShapeRelativeTo.Paragraph, // line not fully supported, map to paragraph
            _ => ShapeRelativeTo.Page
        };
    }

    /// <summary>
    /// Maps FSPA flags to a wrapping mode. For now we conservatively default to
    /// square wrapping, which matches the most common Word behavior for floating
    /// pictures, and fall back to no-wrap only when explicitly requested later.
    /// </summary>
    private static ShapeWrapType MapWrapType(ushort flags)
    {
        // bits 5-7: wr (0=none (behind/ahead), 1=tight, 2=through, 3=square, 4=none (above))
        int wr = (flags >> 5) & 0x07;
        bool fBelowText = (flags & 0x2000) != 0; // bit 13

        return wr switch
        {
            0 => fBelowText ? ShapeWrapType.BehindText : ShapeWrapType.InFrontOfText,
            1 => ShapeWrapType.Tight,
            2 => ShapeWrapType.Through,
            3 => ShapeWrapType.Square,
            4 => ShapeWrapType.TopBottom,
            _ => ShapeWrapType.Square
        };
    }
}

