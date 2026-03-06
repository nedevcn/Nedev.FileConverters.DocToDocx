using System.Runtime.InteropServices;
using System.Text;

namespace Nedev.FileConverters.DocToDocx.Utils;

/// <summary>
/// Extension methods for BinaryReader to handle Word document binary format
/// </summary>
public static class BinaryReaderExt
{
    /// <summary>
    /// Reads a 16-bit big-endian integer
    /// </summary>
    public static short ReadInt16BE(this BinaryReader reader)
    {
        var bytes = reader.ReadBytes(2);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToInt16(bytes, 0);
    }

    /// <summary>
    /// Reads a 16-bit unsigned big-endian integer
    /// </summary>
    public static ushort ReadUInt16BE(this BinaryReader reader)
    {
        var bytes = reader.ReadBytes(2);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToUInt16(bytes, 0);
    }

    /// <summary>
    /// Reads a 32-bit big-endian integer
    /// </summary>
    public static int ReadInt32BE(this BinaryReader reader)
    {
        var bytes = reader.ReadBytes(4);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToInt32(bytes, 0);
    }

    /// <summary>
    /// Reads a 32-bit unsigned big-endian integer
    /// </summary>
    public static uint ReadUInt32BE(this BinaryReader reader)
    {
        var bytes = reader.ReadBytes(4);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToUInt32(bytes, 0);
    }

    /// <summary>
    /// Reads a 64-bit big-endian integer
    /// </summary>
    public static long ReadInt64BE(this BinaryReader reader)
    {
        var bytes = reader.ReadBytes(8);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToInt64(bytes, 0);
    }

    /// <summary>
    /// Reads a 64-bit unsigned big-endian integer
    /// </summary>
    public static ulong ReadUInt64BE(this BinaryReader reader)
    {
        var bytes = reader.ReadBytes(8);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        return BitConverter.ToUInt64(bytes, 0);
    }

    /// <summary>
    /// Reads a Word 6/95 style 16-bit integer (little-endian)
    /// </summary>
    public static short ReadInt16LE(this BinaryReader reader)
    {
        return reader.ReadInt16();
    }

    /// <summary>
    /// Reads a Word 6/95 style 32-bit integer (little-endian)
    /// </summary>
    public static int ReadInt32LE(this BinaryReader reader)
    {
        return reader.ReadInt32();
    }

    /// <summary>
    /// Reads a COMSTAT structure for FILETIME (Windows FILETIME - 64-bit)
    /// </summary>
    public static long ReadFileTime(this BinaryReader reader)
    {
        return reader.ReadInt64();
    }

    /// <summary>
    /// Reads a Word-style date/time (seconds since 1899-12-30)
    /// </summary>
    public static DateTime ReadWordDateTime(this BinaryReader reader)
    {
        var days = reader.ReadDouble();
        // Excel/Word uses 1900-based date system
        // Day 1 = 1900-01-01 (actually 1899-12-31 due to Excel's bug)
        return new DateTime(1899, 12, 30).AddDays(days);
    }

    /// <summary>
    /// Reads a null-terminated Unicode string (UTF-16LE)
    /// </summary>
    public static string ReadUnicodeString(this BinaryReader reader, int maxLength = 256)
    {
        var chars = new List<char>();
        int count = 0;
        while (count < maxLength)
        {
            var ch = reader.ReadChar();
            if (ch == '\0') break;
            chars.Add(ch);
            count++;
        }
        return new string(chars.ToArray());
    }

    /// <summary>
    /// Reads a null-terminated ANSI string
    /// </summary>
    public static string ReadAnsiString(this BinaryReader reader, int maxLength = 256)
    {
        var bytes = new List<byte>();
        int count = 0;
        while (count < maxLength)
        {
            var b = reader.ReadByte();
            if (b == 0) break;
            bytes.Add(b);
            count++;
        }
        return Encoding.ASCII.GetString(bytes.ToArray());
    }

    /// <summary>
    /// Reads a Pascal-style string (length byte + characters)
    /// </summary>
    public static string ReadPascalString(this BinaryReader reader, Encoding? encoding = null)
    {
        encoding ??= Encoding.ASCII;
        var length = reader.ReadByte();
        if (length == 0) return string.Empty;
        var bytes = reader.ReadBytes(length);
        return encoding.GetString(bytes);
    }

    /// <summary>
    /// Reads a count-prefixed string (16-bit length + characters)
    /// </summary>
    public static string ReadCountedString16(this BinaryReader reader, Encoding? encoding = null)
    {
        encoding ??= Encoding.Unicode;
        var length = reader.ReadUInt16();
        if (length == 0) return string.Empty;
        var bytes = reader.ReadBytes(length * 2);
        return encoding.GetString(bytes);
    }

    /// <summary>
    /// Reads a count-prefixed string (32-bit length + characters)
    /// </summary>
    public static string ReadCountedString32(this BinaryReader reader, Encoding? encoding = null)
    {
        encoding ??= Encoding.Unicode;
        var length = reader.ReadInt32();
        if (length == 0) return string.Empty;
        if (length < 0 || length > 10000) return string.Empty; // Sanity check
        var bytes = reader.ReadBytes(length * 2);
        return encoding.GetString(bytes);
    }

    /// <summary>
    /// Skips bytes in the stream
    /// </summary>
    public static void Skip(this BinaryReader reader, int count)
    {
        reader.BaseStream.Seek(count, SeekOrigin.Current);
    }

    /// <summary>
    /// Reads exactly the specified number of bytes, throwing if not available
    /// </summary>
    public static byte[] ReadExact(this BinaryReader reader, int count)
    {
        var result = reader.ReadBytes(count);
        if (result.Length != count)
            throw new EndOfStreamException($"Expected {count} bytes but got {result.Length}");
        return result;
    }

    /// <summary>
    /// Gets current position in stream
    /// </summary>
    public static long Position(this BinaryReader reader)
    {
        return reader.BaseStream.Position;
    }

    /// <summary>
    /// Sets position in stream
    /// </summary>
    public static void Seek(this BinaryReader reader, long position)
    {
        reader.BaseStream.Seek(position, SeekOrigin.Begin);
    }

    /// <summary>
    /// Gets remaining bytes available
    /// </summary>
    public static long Remaining(this BinaryReader reader)
    {
        return reader.BaseStream.Length - reader.BaseStream.Position;
    }

    /// <summary>
    /// Reads 24-bit integer (3 bytes)
    /// </summary>
    public static int ReadInt24(this BinaryReader reader)
    {
        var bytes = reader.ReadBytes(3);
        return bytes[0] | (bytes[1] << 8) | (bytes[2] << 16);
    }

    /// <summary>
    /// Reads unsigned 24-bit integer (3 bytes)
    /// </summary>
    public static uint ReadUInt24(this BinaryReader reader)
    {
        return (uint)ReadInt24(reader);
    }

    /// <summary>
    /// Reads a fixed-size buffer
    /// </summary>
    public static void ReadFixedBuffer(this BinaryReader reader, byte[] buffer)
    {
        var read = reader.Read(buffer, 0, buffer.Length);
        if (read != buffer.Length)
            throw new EndOfStreamException($"Expected {buffer.Length} bytes but got {read}");
    }

    /// <summary>
    /// Creates a slice of the stream for reading a portion
    /// </summary>
    public static BinaryReader CreateSlice(this BinaryReader reader, long offset, int length)
    {
        var baseStream = reader.BaseStream;
        var originalPosition = baseStream.Position;
        
        var sliceStream = new MemoryStream();
        baseStream.Seek(offset, SeekOrigin.Begin);
        var data = baseStream.Length >= offset + length 
            ? reader.ReadBytes(length) 
            : reader.ReadBytes((int)(baseStream.Length - offset));
        sliceStream.Write(data, 0, data.Length);
        sliceStream.Position = 0;
        
        baseStream.Seek(originalPosition, SeekOrigin.Begin);
        return new BinaryReader(sliceStream);
    }
}

/// <summary>
/// Fast binary writer for output operations
/// </summary>
public static class BinaryWriterExt
{
    /// <summary>
    /// Writes a 16-bit big-endian integer
    /// </summary>
    public static void WriteInt16BE(this BinaryWriter writer, short value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        writer.Write(bytes);
    }

    /// <summary>
    /// Writes a 16-bit unsigned big-endian integer
    /// </summary>
    public static void WriteUInt16BE(this BinaryWriter writer, ushort value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        writer.Write(bytes);
    }

    /// <summary>
    /// Writes a 32-bit big-endian integer
    /// </summary>
    public static void WriteInt32BE(this BinaryWriter writer, int value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        writer.Write(bytes);
    }

    /// <summary>
    /// Writes a 32-bit unsigned big-endian integer
    /// </summary>
    public static void WriteUInt32BE(this BinaryWriter writer, uint value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        writer.Write(bytes);
    }

    /// <summary>
    /// Writes a 64-bit big-endian integer
    /// </summary>
    public static void WriteInt64BE(this BinaryWriter writer, long value)
    {
        var bytes = BitConverter.GetBytes(value);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(bytes);
        writer.Write(bytes);
    }

    /// <summary>
    /// Writes a null-terminated Unicode string
    /// </summary>
    public static void WriteUnicodeString(this BinaryWriter writer, string value)
    {
        foreach (var ch in value)
        {
            writer.Write(ch);
        }
        writer.Write('\0');
    }

    /// <summary>
    /// Writes a null-terminated ANSI string
    /// </summary>
    public static void WriteAnsiString(this BinaryWriter writer, string value)
    {
        var bytes = Encoding.ASCII.GetBytes(value);
        writer.Write(bytes);
        writer.Write((byte)0);
    }

    /// <summary>
    /// Writes a Pascal-style string (length byte + characters)
    /// </summary>
    public static void WritePascalString(this BinaryWriter writer, string value, Encoding? encoding = null)
    {
        encoding ??= Encoding.ASCII;
        var bytes = encoding.GetBytes(value);
        if (bytes.Length > 255) bytes = bytes[..255];
        writer.Write((byte)bytes.Length);
        writer.Write(bytes);
    }
}
