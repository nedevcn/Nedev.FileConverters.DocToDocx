using System.Text;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// OLE2 / Compound File Binary (MS-CFB) container reader.
/// Parses the CFB structure to extract named streams such as
/// "WordDocument", "1Table"/"0Table", and "Data" from .doc files.
/// </summary>
public partial class CfbReader : IDisposable
{
    // CFB magic signature: D0 CF 11 E0 A1 B1 1A E1
    private static readonly byte[] CfbSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    // Special FAT sector values
    private const uint DIFSECT   = 0xFFFFFFFC; // DIFAT sector
    private const uint FATSECT   = 0xFFFFFFFD; // FAT sector
    private const uint ENDOFCHAIN = 0xFFFFFFFE; // End of chain
    private const uint FREESECT  = 0xFFFFFFFF; // Free sector

    // Special directory entry stream IDs
    private const uint NOSTREAM = 0xFFFFFFFF;

    // Directory entry object types
    private const byte OBJ_UNKNOWN   = 0x00;
    private const byte OBJ_STORAGE   = 0x01;
    private const byte OBJ_STREAM    = 0x02;
    private const byte OBJ_ROOT      = 0x05;

    // Header fields
    private ushort _majorVersion;
    private ushort _minorVersion;
    private int _sectorSize;       // 512 (V3) or 4096 (V4)
    private int _miniSectorSize;   // always 64
    private int _sectorShift;
    private int _miniSectorShift;
    private uint _totalFatSectors;
    private uint _firstDirectorySectorId;
    private uint _miniStreamCutoff; // 4096
    private uint _firstMiniFatSectorId;
    private uint _totalMiniFatSectors;
    private uint _firstDifatSectorId;
    private uint _totalDifatSectors;

    // Parsed data
    private uint[] _fat = Array.Empty<uint>();
    private uint[] _miniFat = Array.Empty<uint>();
    private List<DirectoryEntry> _directory = new();
    private byte[] _miniStream = Array.Empty<byte>();

    private readonly Stream _stream;
    private readonly BinaryReader _reader;
    private readonly bool _leaveOpen;
    private uint _encryptionKey = 0;

    /// <summary>
    /// All directory entries in the compound file
    /// </summary>
    public IReadOnlyList<DirectoryEntry> Directory => _directory;

    /// <summary>
    /// Names of all streams available in the compound file
    /// </summary>
    public IReadOnlyList<string> StreamNames =>
        _directory.Where(d => d.ObjectType == OBJ_STREAM)
                  .Select(d => d.Name)
                  .ToList();

    public CfbReader(string filePath)
    {
        _stream = File.OpenRead(filePath);
        _reader = new BinaryReader(_stream, Encoding.Default, leaveOpen: false);
        _leaveOpen = false;
        Parse();
    }

    public CfbReader(Stream stream, bool leaveOpen = false)
    {
        _stream = stream;
        _reader = new BinaryReader(_stream, Encoding.Default, leaveOpen: true);
        _leaveOpen = leaveOpen;
        Parse();
    }

    /// <summary>
    /// Checks whether a named stream exists in the compound file
    /// </summary>
    public bool HasStream(string name)
    {
        return _directory.Any(d => d.ObjectType == OBJ_STREAM &&
                                    d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Extracts a named stream as a MemoryStream.
    /// Throws KeyNotFoundException if the stream does not exist.
    /// </summary>
    public MemoryStream GetStream(string name)
    {
        var entry = _directory.FirstOrDefault(d =>
            (d.ObjectType == OBJ_STREAM || d.ObjectType == OBJ_ROOT) &&
            d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

        if (entry == null)
            throw new KeyNotFoundException($"Stream '{name}' not found in compound file. Available: {string.Join(", ", StreamNames)}");

        return ReadStream(entry);
    }

    /// <summary>
    /// Extracts a named stream as a byte array.
    /// </summary>
    public byte[] GetStreamBytes(string name)
    {
        using var ms = GetStream(name);
        return ms.ToArray();
    }

    /// <summary>
    /// Gets a DirectoryEntry by looking up a path like "ObjectPool/_12345678".
    /// </summary>
    public DirectoryEntry? GetStorage(string path)
    {
        if (_directory.Count == 0) return null;
        var parts = path.Split('/');
        var current = _directory[0]; // Root

        foreach (var part in parts)
        {
            current = FindChild(current, part);
            if (current == null) return null;
        }

        return current;
    }

    private DirectoryEntry? FindChild(DirectoryEntry parent, string name)
    {
        if (parent.ChildSid == NOSTREAM) return null;
        return FindNode(parent.ChildSid, name);
    }

    private DirectoryEntry? FindNode(uint sid, string name)
    {
        if (sid == NOSTREAM || sid >= _directory.Count) return null;
        var node = _directory[(int)sid];
        
        // CFB strings are case-insensitive
        if (string.Equals(name, node.Name, StringComparison.OrdinalIgnoreCase)) return node;
        
        var left = FindNode(node.LeftChildSid, name);
        if (left != null) return left;
        
        return FindNode(node.RightChildSid, name);
    }

    /// <summary>
    /// Checks if a child storage exists.
    /// </summary>
    public bool HasStorage(string name)
    {
        if (_directory.Count == 0) return false;
        var node = FindChild(_directory[0], name);
        return node != null && node.ObjectType == 1; // 1 = Storage
    }

    /// <summary>
    /// Gets all immediate children of a storage DirectoryEntry.
    /// </summary>
    public IEnumerable<DirectoryEntry> GetChildren(DirectoryEntry parent)
    {
        var children = new List<DirectoryEntry>();
        if (parent.ChildSid != NOSTREAM)
        {
            CollectNodes(parent.ChildSid, children);
        }
        return children;
    }

    private void CollectNodes(uint sid, List<DirectoryEntry> nodes)
    {
        if (sid == NOSTREAM || sid >= _directory.Count) return;
        var node = _directory[(int)sid];
        nodes.Add(node);
        CollectNodes(node.LeftChildSid, nodes);
        CollectNodes(node.RightChildSid, nodes);
    }

    /// <summary>
    /// Extracts a stream as a byte array given its DirectoryEntry.
    /// </summary>
    public byte[] GetStreamBytes(DirectoryEntry entry)
    {
        using var ms = ReadStream(entry);
        return ms.ToArray();
    }

    // ─── Private: Full parse pipeline ───────────────────────────────

    private void Parse()
    {
        ReadHeader();
        ReadFat();
        ReadDirectory();
        ReadMiniFat();
        LoadMiniStream();
    }

    // ─── Header (512 bytes at offset 0) ─────────────────────────────

    private void ReadHeader()
    {
        _stream.Seek(0, SeekOrigin.Begin);

        // Signature (8 bytes)
        var sig = _reader.ReadBytes(8);
        if (!sig.SequenceEqual(CfbSignature))
            throw new InvalidDataException(
                $"Not a valid OLE2/CFB file. Expected signature D0CF11E0A1B11AE1, got {BitConverter.ToString(sig)}");

        // CLSID (16 bytes) — skip
        _reader.ReadBytes(16);

        // Minor version (2 bytes)
        _minorVersion = _reader.ReadUInt16();

        // Major version (2 bytes): 3 = V3 (512-byte sectors), 4 = V4 (4096-byte sectors)
        _majorVersion = _reader.ReadUInt16();
        if (_majorVersion != 3 && _majorVersion != 4)
            throw new InvalidDataException($"Unsupported CFB major version: {_majorVersion}");

        // Byte order (2 bytes) — must be 0xFFFE (little-endian)
        var byteOrder = _reader.ReadUInt16();
        if (byteOrder != 0xFFFE)
            throw new InvalidDataException($"Unsupported byte order: 0x{byteOrder:X4}");

        // Sector size power (2 bytes)
        _sectorShift = _reader.ReadUInt16();
        _sectorSize = 1 << _sectorShift;

        // Mini sector size power (2 bytes)
        _miniSectorShift = _reader.ReadUInt16();
        _miniSectorSize = 1 << _miniSectorShift;

        // Reserved (6 bytes)
        _reader.ReadBytes(6);

        // Total directory sectors (4 bytes) — V3 must be 0, V4 is actual count
        _reader.ReadUInt32(); // we derive directory from FAT chain

        // Total FAT sectors (4 bytes)
        _totalFatSectors = _reader.ReadUInt32();

        // First directory sector SID (4 bytes)
        _firstDirectorySectorId = _reader.ReadUInt32();

        // Transaction signature (4 bytes) — skip
        _reader.ReadUInt32();

        // Mini stream cutoff size (4 bytes) — typically 4096
        _miniStreamCutoff = _reader.ReadUInt32();

        // First mini FAT sector SID (4 bytes)
        _firstMiniFatSectorId = _reader.ReadUInt32();

        // Total mini FAT sectors (4 bytes)
        _totalMiniFatSectors = _reader.ReadUInt32();

        // First DIFAT sector SID (4 bytes)
        _firstDifatSectorId = _reader.ReadUInt32();

        // Total DIFAT sectors (4 bytes)
        _totalDifatSectors = _reader.ReadUInt32();

        // DIFAT array in header (109 entries × 4 bytes = 436 bytes)
        // We read these as part of ReadFat()
    }

    // ─── FAT (File Allocation Table) ────────────────────────────────

    private void ReadFat()
    {
        // Step 1: collect all FAT sector IDs from the header DIFAT (109 entries)
        //         and any additional DIFAT sectors chained from first DIFAT sector.
        var fatSectorIds = new List<uint>();

        // Header contains 109 DIFAT entries starting at offset 76
        _stream.Seek(76, SeekOrigin.Begin);
        for (int i = 0; i < 109; i++)
        {
            var sid = _reader.ReadUInt32();
            if (sid != FREESECT && sid != ENDOFCHAIN)
                fatSectorIds.Add(sid);
        }

        // Follow DIFAT chain if more than 109 FAT sectors
        if (_totalDifatSectors > 0 && _firstDifatSectorId != ENDOFCHAIN)
        {
            var difatSid = _firstDifatSectorId;
            for (uint d = 0; d < _totalDifatSectors && difatSid != ENDOFCHAIN && difatSid != FREESECT; d++)
            {
                var difatData = ReadRawSector(difatSid);
                var entriesPerSector = (_sectorSize / 4) - 1; // last entry is next DIFAT SID
                for (int i = 0; i < entriesPerSector; i++)
                {
                    var sid = BitConverter.ToUInt32(difatData, i * 4);
                    if (sid != FREESECT && sid != ENDOFCHAIN)
                        fatSectorIds.Add(sid);
                }
                // Last 4 bytes = next DIFAT sector
                difatSid = BitConverter.ToUInt32(difatData, entriesPerSector * 4);
            }
        }

        // Step 2: read all FAT sectors and concatenate into one big FAT array
        var entriesPerFatSector = _sectorSize / 4;
        _fat = new uint[fatSectorIds.Count * entriesPerFatSector];

        for (int s = 0; s < fatSectorIds.Count; s++)
        {
            var sectorData = ReadRawSector(fatSectorIds[s]);
            for (int i = 0; i < entriesPerFatSector; i++)
            {
                _fat[s * entriesPerFatSector + i] = BitConverter.ToUInt32(sectorData, i * 4);
            }
        }
    }

    // ─── Directory ──────────────────────────────────────────────────

    private void ReadDirectory()
    {
        _directory = new List<DirectoryEntry>();

        // Read all directory sectors following the FAT chain
        var dirData = ReadFatChain(_firstDirectorySectorId);
        if (dirData.Length == 0) return;

        // Each directory entry is 128 bytes
        var entryCount = dirData.Length / 128;
        for (int i = 0; i < entryCount; i++)
        {
            var offset = i * 128;
            var entry = ParseDirectoryEntry(dirData, offset, i);
            if (entry.ObjectType != OBJ_UNKNOWN)
            {
                _directory.Add(entry);
            }
            else if (i == 0)
            {
                // Even if type=0, entry 0 is always root — shouldn't happen in valid files
                break;
            }
        }
    }

    private DirectoryEntry ParseDirectoryEntry(byte[] data, int offset, int index)
    {
        var entry = new DirectoryEntry { Index = index };

        // Name (64 bytes, UTF-16LE, null-terminated)
        var nameSize = BitConverter.ToUInt16(data, offset + 64); // in bytes including null
        if (nameSize > 64) nameSize = 64;
        if (nameSize >= 2)
        {
            entry.Name = Encoding.Unicode.GetString(data, offset, nameSize - 2); // exclude null terminator
        }
        else
        {
            entry.Name = string.Empty;
        }

        // Object type (1 byte at offset+66)
        entry.ObjectType = data[offset + 66];

        // Color flag (1 byte at offset+67): 0=red, 1=black (for red-black tree)
        entry.ColorFlag = data[offset + 67];

        // Left child SID (4 bytes at offset+68)
        entry.LeftChildSid = BitConverter.ToUInt32(data, offset + 68);

        // Right child SID (4 bytes at offset+72)
        entry.RightChildSid = BitConverter.ToUInt32(data, offset + 72);

        // Child (root of sub-tree) SID (4 bytes at offset+76)
        entry.ChildSid = BitConverter.ToUInt32(data, offset + 76);

        // CLSID (16 bytes at offset+80) — skip

        // State bits (4 bytes at offset+96) — skip

        // Creation time (8 bytes at offset+100) — skip
        // Modification time (8 bytes at offset+108) — skip

        // Starting sector SID (4 bytes at offset+116)
        entry.StartSectorId = BitConverter.ToUInt32(data, offset + 116);

        // Stream size (for V3: low 4 bytes at offset+120; for V4: 8 bytes at offset+120)
        if (_majorVersion == 4)
        {
            entry.StreamSize = (long)BitConverter.ToUInt64(data, offset + 120);
        }
        else
        {
            entry.StreamSize = BitConverter.ToUInt32(data, offset + 120);
        }

        return entry;
    }

    // ─── Mini FAT ───────────────────────────────────────────────────

    private void ReadMiniFat()
    {
        if (_firstMiniFatSectorId == ENDOFCHAIN || _totalMiniFatSectors == 0)
        {
            _miniFat = Array.Empty<uint>();
            return;
        }

        var miniFatData = ReadFatChain(_firstMiniFatSectorId);
        var entryCount = miniFatData.Length / 4;
        _miniFat = new uint[entryCount];
        for (int i = 0; i < entryCount; i++)
        {
            _miniFat[i] = BitConverter.ToUInt32(miniFatData, i * 4);
        }
    }

    private void LoadMiniStream()
    {
        // The mini stream is stored as data of the root directory entry
        if (_directory.Count == 0) return;

        var root = _directory[0]; // Entry 0 is always root
        if (root.ObjectType != OBJ_ROOT || root.StreamSize == 0)
        {
            _miniStream = Array.Empty<byte>();
            return;
        }

        // Read the root entry's stream data using FAT chain — this IS the mini stream container
        _miniStream = ReadFatChain(root.StartSectorId);
        // Trim to actual size
        if (_miniStream.Length > root.StreamSize)
        {
            Array.Resize(ref _miniStream, (int)root.StreamSize);
        }
    }

    // ─── Stream reading ─────────────────────────────────────────────

    private MemoryStream ReadStream(DirectoryEntry entry)
    {
        byte[] data;

        if (entry.ObjectType != OBJ_ROOT && entry.StreamSize < _miniStreamCutoff)
        {
            // Small stream — read from mini stream using mini FAT
            data = ReadMiniStream(entry.StartSectorId, (int)entry.StreamSize);
        }
        else
        {
            // Regular stream — read using main FAT
            data = ReadFatChain(entry.StartSectorId);
        }

        // Trim to actual stream size
        var actualSize = (int)Math.Min(entry.StreamSize, data.Length);
        return new MemoryStream(data, 0, actualSize, writable: false);
    }

    /// <summary>
    /// Reads a chain of sectors from the main FAT, returning concatenated data.
    /// </summary>
    private byte[] ReadFatChain(uint startSectorId)
    {
        if (startSectorId == ENDOFCHAIN || startSectorId == FREESECT)
            return Array.Empty<byte>();

        var sectors = new List<byte[]>();
        var sid = startSectorId;
        var visited = new HashSet<uint>(); // cycle detection

        while (sid != ENDOFCHAIN && sid != FREESECT)
        {
            if (!visited.Add(sid))
                throw new InvalidDataException($"Cycle detected in FAT chain at sector {sid}");

            if (sid >= _fat.Length)
                break; // safety

            sectors.Add(ReadRawSector(sid));
            sid = _fat[sid];
        }

        // Concatenate all sector data
        var totalLen = sectors.Sum(s => s.Length);
        var result = new byte[totalLen];
        var pos = 0;
        foreach (var sec in sectors)
        {
            Buffer.BlockCopy(sec, 0, result, pos, sec.Length);
            pos += sec.Length;
        }
        return result;
    }

    /// <summary>
    /// Reads data from the mini stream using the mini FAT chain.
    /// </summary>
    private byte[] ReadMiniStream(uint startMiniSectorId, int streamSize)
    {
        if (startMiniSectorId == ENDOFCHAIN || _miniStream.Length == 0)
            return Array.Empty<byte>();

        var result = new byte[streamSize];
        var pos = 0;
        var sid = startMiniSectorId;
        var visited = new HashSet<uint>();

        while (sid != ENDOFCHAIN && sid != FREESECT && pos < streamSize)
        {
            if (!visited.Add(sid))
                throw new InvalidDataException($"Cycle detected in mini FAT chain at sector {sid}");

            if (sid >= _miniFat.Length)
                break;

            var miniOffset = (int)(sid * _miniSectorSize);
            var bytesToRead = Math.Min(_miniSectorSize, streamSize - pos);

            if (miniOffset + bytesToRead <= _miniStream.Length)
            {
                Buffer.BlockCopy(_miniStream, miniOffset, result, pos, bytesToRead);
            }

            pos += bytesToRead;
            sid = _miniFat[sid];
        }

        return result;
    }

    /// <summary>
    /// Reads a single raw sector from the file.
    /// Sector 0 starts immediately after the 512-byte header.
    /// </summary>
    private byte[] ReadRawSector(uint sectorId)
    {
        // In CFB, sector 0 starts at file offset = sectorSize (for V3, that's 512).
        // The header always occupies the first 512 bytes.
        // For V3: sector N is at offset (N + 1) * 512
        // For V4: header is padded to 4096 bytes, sector N is at offset (N + 1) * 4096
        long fileOffset = (sectorId + 1) * _sectorSize;

        if (fileOffset + _sectorSize > _stream.Length)
        {
            // Return partial sector padded with zeros
            var partial = new byte[_sectorSize];
            _stream.Seek(fileOffset, SeekOrigin.Begin);
            var available = (int)Math.Max(0, _stream.Length - fileOffset);
            if (available > 0)
                _ = _stream.Read(partial, 0, available);
            return partial;
        }

        _stream.Seek(fileOffset, SeekOrigin.Begin);
        var buffer = new byte[_sectorSize];
        var read = _stream.Read(buffer, 0, _sectorSize);
        if (read < _sectorSize)
        {
            // Pad remainder with zeros (already zeroed by default)
        }
        return buffer;
    }

    // ─── Diagnostics ────────────────────────────────────────────────

    /// <summary>
    /// Returns a human-readable summary of the compound file structure.
    /// Useful for debugging.
    /// </summary>
    public string GetDiagnostics()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"CFB Version: {_majorVersion}.{_minorVersion}");
        sb.AppendLine($"Sector size: {_sectorSize}");
        sb.AppendLine($"Mini sector size: {_miniSectorSize}");
        sb.AppendLine($"Mini stream cutoff: {_miniStreamCutoff}");
        sb.AppendLine($"FAT sectors: {_totalFatSectors}, entries: {_fat.Length}");
        sb.AppendLine($"Mini FAT sectors: {_totalMiniFatSectors}, entries: {_miniFat.Length}");
        sb.AppendLine($"Mini stream size: {_miniStream.Length}");
        sb.AppendLine($"Directory entries: {_directory.Count}");
        sb.AppendLine();
        sb.AppendLine("Directory listing:");
        foreach (var entry in _directory)
        {
            var typeStr = entry.ObjectType switch
            {
                OBJ_ROOT => "ROOT",
                OBJ_STORAGE => "STORAGE",
                OBJ_STREAM => "STREAM",
                _ => $"TYPE({entry.ObjectType})"
            };
            sb.AppendLine($"  [{entry.Index}] {typeStr,-8} \"{entry.Name}\" size={entry.StreamSize} start={entry.StartSectorId}");
        }
        return sb.ToString();
    }

    public void Dispose()
    {
        _reader?.Dispose();
        if (!_leaveOpen)
            _stream?.Dispose();
    }
}

/// <summary>
/// Represents a directory entry in a Compound File Binary container.
/// </summary>
public class DirectoryEntry
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public byte ObjectType { get; set; }
    public byte ColorFlag { get; set; }
    public uint LeftChildSid { get; set; }
    public uint RightChildSid { get; set; }
    public uint ChildSid { get; set; }
    public uint StartSectorId { get; set; }
    public long StreamSize { get; set; }
}

/// <summary>
/// Provides methods to work with encrypted streams in CFB files.
/// </summary>
public partial class CfbReader
{
    /// <summary>
    /// Sets the encryption key for XOR-encrypted streams.
    /// </summary>
    /// <param name="key">The XOR encryption key (typically from FIB.LKey).</param>
    public void SetEncryptionKey(uint key)
    {
        _encryptionKey = key;
    }

    /// <summary>
    /// Gets a stream and decrypts it if it's XOR-encrypted.
    /// </summary>
    public MemoryStream GetDecryptedStream(string name)
    {
        var stream = GetStream(name);
        
        if (_encryptionKey != 0)
        {
            var encryptedBytes = stream.ToArray();
            var decryptedBytes = EncryptionHelper.DecryptXor(encryptedBytes, _encryptionKey);
            return new MemoryStream(decryptedBytes);
        }
        
        return stream;
    }
}
