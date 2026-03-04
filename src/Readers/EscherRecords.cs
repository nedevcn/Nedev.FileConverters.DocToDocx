using System.Text;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Base Escher record (MS-ODRAW) as used by OfficeArt.
/// This is a low-level representation that keeps raw data and a tree of children.
/// </summary>
public class EscherRecord
{
    public ushort Version { get; set; }
    public ushort Instance { get; set; }
    public ushort Type { get; set; }
    public uint Length { get; set; }
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public List<EscherRecord> Children { get; set; } = new();

    public bool IsContainer => Version == 0xF;

    public override string ToString()
    {
        var name = $"0x{Type:X4}";
        return $"{name} ver={Version} inst={Instance} len={Length}";
    }
}

/// <summary>
/// Helper responsible for parsing Escher records from a binary stream.
/// This is intentionally generic and does not interpret record semantics yet.
/// </summary>
public static class EscherReader
{
    public static List<EscherRecord> ReadAll(BinaryReader reader, long maxBytes)
    {
        var records = new List<EscherRecord>();
        var start = reader.BaseStream.Position;

        while (reader.BaseStream.Position - start + 8 <= maxBytes)
        {
            try
            {
                var record = ReadRecord(reader);
                if (record == null)
                {
                    break;
                }
                records.Add(record);
            }
            catch
            {
                // Stop on any parse error; this is best-effort only.
                break;
            }
        }

        return records;
    }

    public static EscherRecord? ReadRecord(BinaryReader reader)
    {
        if (reader.BaseStream.Position + 8 > reader.BaseStream.Length)
            return null;

        var header = reader.ReadUInt16();
        var ver = (ushort)(header & 0x000F);
        var inst = (ushort)((header & 0xFFF0) >> 4);
        var type = reader.ReadUInt16();
        var length = reader.ReadUInt32();

        if (length > int.MaxValue)
            return null;
        if (reader.BaseStream.Position + length > reader.BaseStream.Length)
            return null;

        var record = new EscherRecord
        {
            Version = ver,
            Instance = inst,
            Type = type,
            Length = length
        };

        if (record.IsContainer)
        {
            var endPos = reader.BaseStream.Position + length;
            while (reader.BaseStream.Position + 8 <= endPos)
            {
                var child = ReadRecord(reader);
                if (child == null) break;
                record.Children.Add(child);
            }

            // Skip any leftover bytes if children did not consume the whole payload.
            if (reader.BaseStream.Position < endPos)
            {
                reader.BaseStream.Seek(endPos, SeekOrigin.Begin);
            }
        }
        else
        {
            record.Data = reader.ReadBytes((int)length);
        }

        return record;
    }
}

/// <summary>
/// Entry point for reading OfficeArt (Escher) data from the Word "Data" stream.
/// 当前阶段只负责将 Escher 记录树解析出来，后续阶段再映射为 ShapeModel。
/// </summary>
public class OfficeArtReader
{
    public List<EscherRecord> RootRecords { get; } = new();

    public OfficeArtReader(Stream dataStream)
    {
        if (dataStream == null || !dataStream.CanRead) return;

        // Work on a copy so we don't disturb existing readers.
        using var ms = new MemoryStream();
        dataStream.Position = 0;
        dataStream.CopyTo(ms);
        ms.Position = 0;

        using var br = new BinaryReader(ms, Encoding.Default, leaveOpen: true);
        var maxBytes = ms.Length;
        RootRecords = EscherReader.ReadAll(br, maxBytes);
    }
}

