#nullable enable
using System;
using System.IO;
using System.Linq;
using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class OfficeArtMapperTests
{
    [Fact]
    public void OfficeArtReader_ResyncsPastLeadingBytes_AndPreservesWordArtText()
    {
        byte[] textBytes = Encoding.Unicode.GetBytes("艺术字\0");
        byte[] optPayload = BuildOptPayload(192, textBytes);
        byte[] spRecord = BuildLeafRecord(0xF00A, 0x00CA, BitConverter.GetBytes(42).Concat(new byte[4]).ToArray(), version: 0x2);
        byte[] optRecord = BuildLeafRecord(0xF00B, 1, optPayload, version: 0x3);
        byte[] spContainer = BuildContainerRecord(0xF004, 0, spRecord.Concat(optRecord).ToArray());
        byte[] data = new byte[] { 0x01, 0x02, 0x03, 0x04 }.Concat(spContainer).ToArray();

        using var stream = new MemoryStream(data);
        var reader = new OfficeArtReader(stream);
        var document = new DocumentModel();

        OfficeArtMapper.AttachShapes(document, reader, null);

        var shape = Assert.Single(document.Shapes);
        Assert.Equal(42, shape.Id);
        Assert.Equal(ShapeType.Textbox, shape.Type);
        Assert.Equal("艺术字", shape.Text);
    }

    [Fact]
    public void SampleTextDoc_DoesNotExposeWordArtThroughOfficeArtStreams()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var officeArtReader = (OfficeArtReader?)typeof(DocReader)
            .GetField("_officeArtReader", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
            .GetValue(docReader);
        var fspaAnchors = (System.Collections.ICollection?)typeof(DocReader)
            .GetField("_fspaAnchors", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
            .GetValue(docReader);

        Assert.Empty(docReader.Document.Shapes);
        Assert.Empty(docReader.Document.Textboxes);
        Assert.Equal(0, officeArtReader?.RootRecords.Count ?? 0);
        Assert.Equal(0, fspaAnchors?.Count ?? 0);
    }

    private static byte[] BuildOptPayload(ushort propId, byte[] complexData)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write((ushort)(0x8000 | propId));
        writer.Write((uint)complexData.Length);
        writer.Write(complexData);
        writer.Flush();
        return ms.ToArray();
    }

    private static byte[] BuildLeafRecord(ushort type, ushort instance, byte[] payload, ushort version)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write((ushort)((instance << 4) | version));
        writer.Write(type);
        writer.Write((uint)payload.Length);
        writer.Write(payload);
        writer.Flush();
        return ms.ToArray();
    }

    private static byte[] BuildContainerRecord(ushort type, ushort instance, byte[] children)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write((ushort)((instance << 4) | 0x000F));
        writer.Write(type);
        writer.Write((uint)children.Length);
        writer.Write(children);
        writer.Flush();
        return ms.ToArray();
    }
}