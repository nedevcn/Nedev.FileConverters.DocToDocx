using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Reads FSPA (floating shape) anchors from the PlcfSpaMom structure in the Table stream.
/// This is a best-effort implementation that decodes the shape id (spid) and bounding box.
/// </summary>
public static class FspaReader
{
    private const int FspaSize = 26; // Size of FSPA structure in bytes (per MS-DOC PlcfSpa)

    public static List<FspaInfo> ReadPlcSpaMom(BinaryReader tableReader, FibReader fib)
    {
        var anchors = new List<FspaInfo>();

        if (fib.FcPlcSpaMom == 0 || fib.LcbPlcSpaMom < 4 + FspaSize)
            return anchors;

        try
        {
            tableReader.BaseStream.Seek(fib.FcPlcSpaMom, SeekOrigin.Begin);
            var lcb = (int)fib.LcbPlcSpaMom;

            // PlcfSpa layout: CP[0..n] (n+1 x 4 bytes) + FSPA[0..n-1] (n x 26 bytes)
            // Total size = 4 + n * (4 + 26) => n = (lcb - 4) / 30
            var n = (lcb - 4) / (4 + FspaSize);
            if (n <= 0)
                return anchors;

            var cpArray = new int[n + 1];
            for (int i = 0; i <= n; i++)
            {
                if (tableReader.BaseStream.Position + 4 > tableReader.BaseStream.Length)
                    return anchors;
                cpArray[i] = tableReader.ReadInt32();
            }

            for (int i = 0; i < n; i++)
            {
                if (tableReader.BaseStream.Position + FspaSize > tableReader.BaseStream.Length)
                    break;

                var cp = cpArray[i];

                // FSPA structure: we mainly rely on the first 5 DWORDs (spid + bounding box),
                // but we also keep a copy of the trailing flags so that higher layers can
                // decide how to interpret relative anchors. This keeps parsing robust even
                // if the tail layout changes slightly between Word versions.
                var spid = tableReader.ReadInt32();
                var xaLeft = tableReader.ReadInt32();
                var yaTop = tableReader.ReadInt32();
                var xaRight = tableReader.ReadInt32();
                var yaBottom = tableReader.ReadInt32();

                ushort flags = 0;
                var remaining = FspaSize - (5 * 4);
                if (remaining >= 4)
                {
                    // Read two WORDs: first is often reserved, second contains flags
                    tableReader.ReadUInt16(); // reserved / unused
                    flags = tableReader.ReadUInt16();
                    remaining -= 4;
                }
                if (remaining > 0)
                {
                    tableReader.BaseStream.Seek(remaining, SeekOrigin.Current);
                }

                anchors.Add(new FspaInfo
                {
                    Spid = spid,
                    XaLeft = xaLeft,
                    YaTop = yaTop,
                    XaRight = xaRight,
                    YaBottom = yaBottom,
                    Cp = cp,
                    Flags = flags
                });
            }
        }
        catch
        {
            // Return whatever we managed to parse; caller treats this as best-effort.
        }

        return anchors;
    }
}

/// <summary>
/// Minimal representation of an FSPA anchor for a floating shape.
/// Coordinates are in twips relative to the page/column (per MS-DOC).
/// </summary>
public class FspaInfo
{
    public int Spid { get; set; }
    public int XaLeft { get; set; }
    public int YaTop { get; set; }
    public int XaRight { get; set; }
    public int YaBottom { get; set; }
    public int Cp { get; set; }
    /// <summary>
    /// Raw FSPA flags as stored in the binary document. The exact bit semantics
    /// are interpreted at a higher layer; unknown combinations fall back to
    /// page-relative anchors for safety.
    /// </summary>
    public ushort Flags { get; set; }
}

