using System.IO;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Reads the File Information Block (FIB) from a Word 97-2003 binary document.
/// The FIB contains the main document metadata and pointers (offsets and lengths)
/// to all other structures in the document.
/// </summary>
public class FibReader
{
    private readonly BinaryReader _reader;

    // ── FibBase fields ──────────────────────────────────────────────
    public ushort WIdent { get; private set; }
    public ushort NFib { get; private set; }
    public ushort Unused { get; private set; }
    public ushort Lid { get; private set; }
    public ushort PnNext { get; private set; }
    public bool FDot { get; private set; }
    public bool FGlsy { get; private set; }
    public bool FComplex { get; private set; }
    public bool FHasPic { get; private set; }
    public ushort CQuickSaves { get; private set; }
    public bool FEncrypted { get; private set; }
    public bool FWhichTblStm { get; private set; }
    public bool FExtChar { get; private set; }
    public bool FObfuscated { get; private set; }
    public ushort NFibBack { get; private set; }
    public uint LKey { get; private set; }
    public byte Envr { get; private set; }
    public bool FMac { get; private set; }
    public bool FEmptySpecial { get; private set; }
    public bool FLoadOverridePage { get; private set; }
    public uint FcMin { get; private set; }
    public uint FcMac { get; private set; }

    // ── FibRgLw fields (CP counts) ──────────────────────────────────
    public int CcpText { get; private set; }
    public int CcpFtn { get; private set; }
    public int CcpHdd { get; private set; }
    public int CcpAtn { get; private set; }
    public int CcpEdn { get; private set; }
    public int CcpTxbx { get; private set; }
    public int CcpHdrTxbx { get; private set; }

    // ── FibRgFcLcb fields (Offsets and Lengths) ─────────────────────
    public uint FcStshf { get; private set; }
    public uint LcbStshf { get; private set; }
    public uint FcPlcfSed { get; private set; }
    public uint LcbPlcfSed { get; private set; }
    public uint FcPlcfBteChpx { get; private set; }
    public uint LcbPlcfBteChpx { get; private set; }
    public uint FcPlcfBtePapx { get; private set; }
    public uint LcbPlcfBtePapx { get; private set; }
    public uint FcPlcfFldMom { get; private set; }
    public uint LcbPlcfFldMom { get; private set; }
    public uint FcPlcffndRef { get; private set; }
    public uint LcbPlcffndRef { get; private set; }
    public uint FcPlcffndTxt { get; private set; }
    public uint LcbPlcffndTxt { get; private set; }
    public uint FcPlcfandRef { get; private set; }
    public uint LcbPlcfandRef { get; private set; }
    public uint FcPlcfandTxt { get; private set; }
    public uint LcbPlcfandTxt { get; private set; }
    public uint FcPlcfBkf { get; private set; }
    public uint LcbPlcfBkf { get; private set; }
    public uint FcPlcfBkl { get; private set; }
    public uint LcbPlcfBkl { get; private set; }
    public uint FcSttbfAtnMod { get; private set; }
    public uint LcbSttbfAtnMod { get; private set; }
    public uint FcPlcfAtnbkf { get; private set; }
    public uint LcbPlcfAtnbkf { get; private set; }
    public uint FcPlcfAtnbkl { get; private set; }
    public uint LcbPlcfAtnbkl { get; private set; }
    public uint FcPlcfFldAtn { get; private set; }
    public uint LcbPlcfFldAtn { get; private set; }
    public uint FcPlcfFldEdn { get; private set; }
    public uint LcbPlcfFldEdn { get; private set; }
    public uint FcPlcfFldFtn { get; private set; }
    public uint LcbPlcfFldFtn { get; private set; }
    public uint FcPlcfFldHdr { get; private set; }
    public uint LcbPlcfFldHdr { get; private set; }
    public uint FcPlcfFldTxbx { get; private set; }
    public uint LcbPlcfFldTxbx { get; private set; }
    public uint FcSttbfBkmk { get; private set; }
    public uint LcbSttbfBkmk { get; private set; }
    public uint FcPlcfHdd { get; private set; }
    public uint LcbPlcfHdd { get; private set; }
    public uint FcClx { get; private set; }
    public uint LcbClx { get; private set; }
    public uint FcPlcSpaMom { get; private set; }
    public uint LcbPlcSpaMom { get; private set; }
    public uint FcPlcfendRef { get; private set; }
    public uint LcbPlcfendRef { get; private set; }
    public uint FcPlcfendTxt { get; private set; }
    public uint LcbPlcfendTxt { get; private set; }
    public uint FcFtn { get; private set; }
    public uint LcbFtn { get; private set; }
    public uint FcEnd { get; private set; }
    public uint LcbEnd { get; private set; }
    public uint FcAnot { get; private set; }
    public uint LcbAnot { get; private set; }
    public uint FcTxbx { get; private set; }
    public uint LcbTxbx { get; private set; }
    public uint FcGlsy { get; private set; }
    public uint LcbGlsy { get; private set; }
    public uint FcData { get; private set; }
    public uint LcbData { get; private set; }
    public uint FcPlcfLst { get; private set; }
    public uint LcbPlcfLst { get; private set; }
    public uint FcPlfLfo { get; private set; }
    public uint LcbPlfLfo { get; private set; }
    public uint FcSttbfFfn { get; private set; }
    public uint LcbSttbfFfn { get; private set; }
    public uint FcDop { get; private set; }
    public uint LcbDop { get; private set; }
    public uint FcSttbfRgtlv { get; private set; }
    public uint LcbSttbfRgtlv { get; private set; }

    // Legacy Aliases for compatibility with older code
    public uint StshOffset => FcStshf;
    public bool IsComplex => FComplex;
    public uint DopOffset => FcDop;
    public uint PnFbpClx => FcClx;
    public uint TextBaseOffset => 0; // In standard implementations, this is often treated as 0 relative to stream

    /// <summary>
    /// Name of the Table stream to use ("0Table" or "1Table")
    /// </summary>
    public string TableStreamName => FWhichTblStm ? "1Table" : "0Table";

    private readonly List<(uint fc, uint lcb)> _rgFcLcb = new();

    public FibReader(BinaryReader reader)
    {
        _reader = reader;
    }

    public void Read()
    {
        _reader.BaseStream.Seek(0, SeekOrigin.Begin);
        ReadFibBase();
        ReadFibRgW();
        ReadFibRgLw();
        ReadFibRgFcLcb();
    }

    private void ReadFibBase()
    {
        WIdent = _reader.ReadUInt16();
        if (WIdent != WordConsts.FIB_MAGIC_NUMBER && WIdent != WordConsts.FIB_MAGIC_NUMBER_OLD)
            throw new InvalidDataException($"Invalid magic: 0x{WIdent:X4}");

        NFib = _reader.ReadUInt16();
        Unused = _reader.ReadUInt16();
        Lid = _reader.ReadUInt16();
        PnNext = _reader.ReadUInt16();

        var flagsA = _reader.ReadUInt16();
        FDot         = (flagsA & 0x01) != 0;
        FGlsy        = (flagsA & 0x02) != 0;
        FComplex     = (flagsA & 0x04) != 0;
        FHasPic      = (flagsA & 0x08) != 0;
        CQuickSaves  = (ushort)((flagsA >> 4) & 0x0F);
        FEncrypted   = (flagsA & 0x100) != 0;
        FWhichTblStm = (flagsA & 0x200) != 0;
        FExtChar     = (flagsA & 0x1000) != 0;
        FObfuscated  = (flagsA & 0x8000) != 0;

        NFibBack = _reader.ReadUInt16();
        LKey = _reader.ReadUInt32();
        Envr = _reader.ReadByte();

        var flagsB = _reader.ReadByte();
        FMac               = (flagsB & 0x01) != 0;
        FEmptySpecial      = (flagsB & 0x02) != 0;
        FLoadOverridePage  = (flagsB & 0x04) != 0;

        _reader.ReadUInt16(); // reserved
        _reader.ReadUInt16(); // reserved
        FcMin = _reader.ReadUInt32();
        FcMac = _reader.ReadUInt32();
    }

    private void ReadFibRgW()
    {
        var csw = _reader.ReadUInt16();
        for (int i = 0; i < csw; i++) _reader.ReadUInt16();
    }

    private void ReadFibRgLw()
    {
        var cslw = _reader.ReadUInt16();
        var rglw = new int[cslw];
        for (int i = 0; i < cslw; i++) rglw[i] = _reader.ReadInt32();

        if (cslw > 3)  CcpText    = rglw[3];
        if (cslw > 4)  CcpFtn     = rglw[4];
        if (cslw > 5)  CcpHdd     = rglw[5];
        if (cslw > 7)  CcpAtn     = rglw[7];
        if (cslw > 8)  CcpEdn     = rglw[8];
        if (cslw > 9)  CcpTxbx    = rglw[9];
        if (cslw > 10) CcpHdrTxbx = rglw[10];
    }

    private void ReadFibRgFcLcb()
    {
        var cbRgFcLcb = _reader.ReadUInt16();
        _rgFcLcb.Clear();
        for (int i = 0; i < cbRgFcLcb; i++)
        {
            var fc = _reader.ReadUInt32();
            var lcb = _reader.ReadUInt32();
            _rgFcLcb.Add((fc, lcb));
        }

        (FcStshf, LcbStshf)             = GetFcLcb(0);
        (FcPlcfFldMom, LcbPlcfFldMom)   = GetFcLcb(4);
        (FcPlcffndRef, LcbPlcffndRef)   = GetFcLcb(8);
        (FcPlcffndTxt, LcbPlcffndTxt)   = GetFcLcb(9);
        (FcPlcfandRef, LcbPlcfandRef)   = GetFcLcb(10);
        (FcPlcfandTxt, LcbPlcfandTxt)   = GetFcLcb(11);
        
        (FcPlcfBteChpx, LcbPlcfBteChpx) = GetFcLcb(12);
        (FcPlcfBtePapx, LcbPlcfBtePapx) = GetFcLcb(13);

        (FcPlcfBkf, LcbPlcfBkf)         = GetFcLcb(15);
        (FcPlcfBkl, LcbPlcfBkl)         = GetFcLcb(16);
        (FcSttbfAtnMod, LcbSttbfAtnMod) = GetFcLcb(17);
        (FcPlcfAtnbkf, LcbPlcfAtnbkf)   = GetFcLcb(18);
        (FcPlcfAtnbkl, LcbPlcfAtnbkl)   = GetFcLcb(19);
        (FcPlcfFldAtn, LcbPlcfFldAtn)   = GetFcLcb(20);
        (FcPlcfFldEdn, LcbPlcfFldEdn)   = GetFcLcb(21);
        (FcPlcfFldFtn, LcbPlcfFldFtn)   = GetFcLcb(22);
        (FcPlcfFldHdr, LcbPlcfFldHdr)   = GetFcLcb(23);
        (FcPlcfFldTxbx, LcbPlcfFldTxbx) = GetFcLcb(24);
        
        (FcSttbfBkmk, LcbSttbfBkmk)     = GetFcLcb(26);
        (FcPlcfHdd, LcbPlcfHdd)         = GetFcLcb(31);
        (FcClx, LcbClx)                 = GetFcLcb(34);
        (FcPlcSpaMom, LcbPlcSpaMom)     = GetFcLcb(37);
        (FcPlcfendRef, LcbPlcfendRef)   = GetFcLcb(39);
        (FcPlcfendTxt, LcbPlcfendTxt)   = GetFcLcb(40);
        (FcFtn, LcbFtn)                 = GetFcLcb(41);
        (FcEnd, LcbEnd)                 = GetFcLcb(42);
        (FcAnot, LcbAnot)               = GetFcLcb(43);
        (FcTxbx, LcbTxbx)               = GetFcLcb(44);
        (FcGlsy, LcbGlsy)               = GetFcLcb(45);
        (FcData, LcbData)               = GetFcLcb(46);
        (FcSttbfFfn, LcbSttbfFfn)       = GetFcLcb(IndexSttbfFfn());
        (FcDop, LcbDop)                 = GetFcLcb(IndexDop());
        (FcSttbfRgtlv, LcbSttbfRgtlv)   = GetFcLcb(47);
        (FcPlcfSed, LcbPlcfSed)         = GetFcLcb(33);

        if (cbRgFcLcb > 89)
        {
            (FcPlcfLst, LcbPlcfLst) = GetFcLcb(88);
            (FcPlfLfo, LcbPlfLfo)   = GetFcLcb(89);
        }
    }

    private int IndexDop() => NFib >= 0x00D9 ? 31 : 19; // Simplified
    private int IndexSttbfFfn() => 21; // Standard for 97+

    public (uint fc, uint lcb) GetFcLcb(int index)
    {
        if (index >= 0 && index < _rgFcLcb.Count) return _rgFcLcb[index];
        return (0, 0);
    }
}
