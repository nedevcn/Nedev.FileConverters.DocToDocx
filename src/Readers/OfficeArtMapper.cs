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

    public static void AttachShapes(DocumentModel document, OfficeArtReader? officeArtReader)
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
}

