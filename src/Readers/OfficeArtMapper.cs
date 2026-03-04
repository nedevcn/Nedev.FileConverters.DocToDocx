using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Readers;

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
    /// EscherSp record and use its shapeId as the ShapeModel.Id。当前阶段还不解析
    /// 复杂的 OfficeArt 属性，只是尝试按顺序将形状与已提取的 Images 对应起来，
    /// 将其视为 Picture 形状。这是一种启发式映射，但在常见场景中通常是合理的。
    /// </summary>
    private static ShapeModel? CreateShapeFromSpContainer(EscherRecord spContainer, DocumentModel document, ref int imageIndexCursor)
    {
        int? shapeId = null;

        foreach (var child in spContainer.Children)
        {
            if (child.Type == RecordTypeSp && child.Data.Length >= 4)
            {
                shapeId = BitConverter.ToInt32(child.Data, 0);
                break;
            }
        }

        if (shapeId == null)
        {
            return null;
        }

        // 启发式：如果还有未消费的 ImageModel，就将该形状视为 Picture，
        // 并与当前游标指向的图片绑定。这样可以在很多实际文档中，把
        // OfficeArt 图片形状与已经抽取的图像数据对齐。
        ShapeType type = ShapeType.Unknown;
        int? imageIndex = null;

        if (imageIndexCursor >= 0 && imageIndexCursor < document.Images.Count)
        {
            type = ShapeType.Picture;
            imageIndex = imageIndexCursor;
            imageIndexCursor++;
        }

        return new ShapeModel
        {
            Id = shapeId.Value,
            Type = type,
            Anchor = null,
            ImageIndex = imageIndex,
            Text = null
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
                VerticalRelativeTo = MapRelativeToVertical(fspa.Flags)
            };

            // Map CP to nearest paragraph by MinCp.
            var cp = fspa.Cp;
            var bestPara = paraInfos.FirstOrDefault(p => p.MinCp != int.MaxValue && p.MinCp >= cp);
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
        // TODO: refine based on full MS-DOC FSPA specification and real-world docs.
        return ShapeRelativeTo.Page;
    }

    /// <summary>
    /// Maps FSPA flags to a vertical reference frame. See comments on
    /// MapRelativeToHorizontal for caveats.
    /// </summary>
    private static ShapeRelativeTo MapRelativeToVertical(ushort flags)
    {
        // TODO: refine based on full MS-DOC FSPA specification and real-world docs.
        return ShapeRelativeTo.Page;
    }
}

