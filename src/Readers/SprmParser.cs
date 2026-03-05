using System.IO;
using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class SprmParser
{
    private readonly BinaryReader _reader;
    private readonly long _endPosition;

    public SprmParser(BinaryReader reader, int length)
    {
        _reader = reader;
        _endPosition = reader.BaseStream.Position + length;
    }

    public void ApplyToChp(byte[] grpprl, ChpBase chp)
    {
        if (grpprl.Length == 0) return;
        using var ms = new MemoryStream(grpprl);
        using var reader = new BinaryReader(ms, Encoding.Default, true);
        while (ms.Position < ms.Length)
        {
            try
            {
                var sprm = ReadSprm(reader);
                ApplyChpSprm(sprm, chp);
            }
            catch (EndOfStreamException)
            {
                break;
            }
        }
    }

    public void ApplyToPap(byte[] grpprl, PapBase pap)
    {
        using var ms = new MemoryStream(grpprl);
        using var reader = new BinaryReader(ms, Encoding.Default, true);
        while (ms.Position < ms.Length)
        {
            try { var sprm = ReadSprm(reader); ApplyPapSprm(sprm, pap); }
            catch (EndOfStreamException) { break; }
        }
    }

    public void ApplyToTap(byte[] grpprl, TapBase tap)
    {
        using var ms = new MemoryStream(grpprl);
        using var reader = new BinaryReader(ms, Encoding.Default, true);
        while (ms.Position < ms.Length)
        {
            try { var sprm = ReadSprm(reader); ApplyTapSprm(sprm, tap); }
            catch (EndOfStreamException) { break; }
        }
    }

    public void ApplyToSep(byte[] grpprl, SepBase sep)
    {
        if (grpprl == null || grpprl.Length == 0) return;
        using var ms = new MemoryStream(grpprl);
        using var reader = new BinaryReader(ms, Encoding.Default, true);
        while (ms.Position < ms.Length)
        {
            try
            {
                var sprm = ReadSprm(reader);
                ApplySepSprm(sprm, sep);
            }
            catch (EndOfStreamException)
            {
                break;
            }
        }
    }

    private Sprm ReadSprm(BinaryReader reader)
    {
        var sprm = new Sprm();
        sprm.Code = reader.ReadUInt16();
        sprm.OperandSize = GetOperandSize(sprm.Code);
        switch (sprm.OperandSize)
        {
            case 1: sprm.Operand = reader.ReadByte(); break;
            case 2: sprm.Operand = reader.ReadUInt16(); break;
            case 4: sprm.Operand = reader.ReadUInt32(); break;
            case 3: sprm.Operand = (uint)(reader.ReadByte() | (reader.ReadByte() << 8) | (reader.ReadByte() << 16)); break;
            default:
                if (sprm.OperandSize == 0xFF) { var varSize = reader.ReadByte(); sprm.VariableOperand = reader.ReadBytes(varSize); }
                break;
        }
        return sprm;
    }

    private int GetOperandSize(ushort sprmCode)
    {
        var spra = (sprmCode >> 10) & 0x07;
        return spra switch { 0 => 1, 1 => 2, 2 => 4, 3 => 2, 4 => 2, 5 => 2, 6 => 4, 7 => 0xFF, _ => 1 };
    }

    private void ApplyChpSprm(Sprm sprm, ChpBase chp)
    {
        // Extract 9-bit operation code from the 16-bit Word 97 sprm Code
        var sprmCode = sprm.Code & 0x01FF;
        var sgc = (sprm.Code >> 13) & 0x07;
        
        // sprmCPicLocation (0x6A03) - picture position in Data stream; sgc can be 2 or 3 in practice
        if (sprm.Code == 0x6A03 && sprm.OperandSize == 4)
        {
            chp.FcPic = (uint)sprm.Operand;
            return;
        }

        // sgc=2 is CHP (Character Properties). Some PAP/TAP might override CHP, so relax strictness
        if (sgc != 2 && sgc != 1) return;

        switch (sprmCode)
        {
            // --- Word 97+ (16-bit) SPRM Opcodes ---
            case 0x35: chp.IsBold = sprm.Operand != 0; break; // sprmCFBold
            case 0x36: chp.IsItalic = sprm.Operand != 0; break; // sprmCFItalic
            case 0x37: chp.IsStrikeThrough = sprm.Operand != 0; break; // sprmCFStrike
            case 0x38: chp.IsOutline = sprm.Operand != 0; break; // sprmCFOutline
            case 0x39: chp.IsShadow = sprm.Operand != 0; break; // sprmCFShadow
            case 0x3A: chp.IsSmallCaps = sprm.Operand != 0; break; // sprmCFSmallCaps
            case 0x3B: chp.IsAllCaps = sprm.Operand != 0; break; // sprmCFCaps
            case 0x3C: chp.IsHidden = sprm.Operand != 0; break; // sprmCFVanish
            case 0x3E: chp.Underline = (byte)sprm.Operand; break; // sprmCKul
            case 0x43: chp.FontSize = (byte)sprm.Operand; break; // sprmCHps (half-points)
            case 0x45: chp.Position = (int)(short)sprm.Operand; break; // sprmCHpsPos
            case 0x4B: chp.Scale = (int)sprm.Operand; break; // sprmCHwcr
            case 0x4F: chp.FontIndex = (short)sprm.Operand; break; // sprmCRqftc
            case 0x5C: chp.IsBoldCs = sprm.Operand != 0; break; // sprmCFBoldBi
            case 0x5D: chp.IsItalicCs = sprm.Operand != 0; break; // sprmCFItalicBi
            case 0x61: chp.FontSizeCs = (byte)sprm.Operand; break; // sprmCHpsBi
            case 0x70: // sprmCCv (24-bit RGB)
                chp.RgbColor = sprm.Operand;
                chp.HasRgbColor = true;
                break;
            case 0x42: chp.Color = (byte)sprm.Operand; break; // sprmCIco
            case 0x0C: chp.HighlightColor = (byte)sprm.Operand; break; // sprmCHighlight
            case 0x68: chp.Language = (int)sprm.Operand; break; // sprmCRgLid0
            case 0x5E: chp.FontIndexCs = (short)sprm.Operand; break; // sprmCRqftcBi
            
            // --- Word 6 (8-bit) SPRM Opcodes (Fallbacks) ---
            case 0x02: chp.IsBold = sprm.Operand != 0; break;
            case 0x03: chp.IsItalic = sprm.Operand != 0; break;
            case 0x04: 
                if (sprm.Code == 0x4804) chp.AuthorIndexDel = (ushort)sprm.Operand;
                else chp.IsStrikeThrough = sprm.Operand != 0; 
                break;
            case 0x05: 
                if (sprm.Code == 0x6805) chp.DateDel = (uint)sprm.Operand;
                else chp.IsUnderline = sprm.Operand != 0; 
                break;
            case 0x06: chp.IsOutline = sprm.Operand != 0; break;
            case 0x07: chp.IsSmallCaps = sprm.Operand != 0; break;
            case 0x08: chp.IsAllCaps = sprm.Operand != 0; break;
            case 0x09: chp.IsHidden = sprm.Operand != 0; break;
            case 0x0A: chp.FontIndex = (short)sprm.Operand; break;
            case 0x0B: chp.Underline = (byte)sprm.Operand; break;
            // case 0x0C: chp.Kerning = (int)(short)sprm.Operand; break; // Conflicts with Word 97 sprmCHighlight
            case 0x0D: chp.Position = (int)(short)sprm.Operand; break;
            case 0x0E: chp.Scale = (int)sprm.Operand; break;
            case 0x11: chp.Color = (byte)sprm.Operand; break;
            case 0x12: chp.FontSize = (byte)sprm.Operand; break;
            case 0x13: chp.HighlightColor = (byte)sprm.Operand; break;
            case 0x16: chp.IsShadow = sprm.Operand != 0; break;
            case 0x17: chp.IsEmboss = sprm.Operand != 0; break;
            case 0x18: chp.IsImprint = sprm.Operand != 0; break;
            case 0x2E: chp.FontIndexCs = (short)sprm.Operand; break;
            case 0x30: chp.FontSizeCs = (byte)sprm.Operand; break;
            case 0x31: chp.IsBoldCs = sprm.Operand != 0; break;
            case 0x32: chp.IsItalicCs = sprm.Operand != 0; break;
            case 0x40: chp.IsSuperscript = sprm.Operand == 1; chp.IsSubscript = sprm.Operand == 2; break;
            case 0x41: chp.IsDoubleStrikeThrough = sprm.Operand != 0; break;
            case 0x44: chp.CharacterSpacingAdjustment = (int)sprm.Operand; break;

            // --- Revision Marks (Track Changes) ---
            case 0x00: chp.IsDeleted = sprm.Operand != 0; break;      // sprmCFRMarkDel
            // sprmCIBstRMarkDel and sprmCDttmRMarkDel logic merged into 0x04 and 0x05 cases above
            case 0x54: chp.IsInserted = sprm.Operand != 0; break;     // sprmCFRMark
            case 0x63: chp.AuthorIndexIns = (ushort)sprm.Operand; break; // sprmCIBstRMark
            case 0x64: chp.DateIns = (uint)sprm.Operand; break;      // sprmCDttmRMark
        }
    }

    private void ApplyPapSprm(Sprm sprm, PapBase pap)
    {
        var sprmCode = sprm.Code & 0x01FF;
        var sgc = (sprm.Code >> 13) & 0x07;
        
        // sgc=1 is PAP (Paragraph Properties), but we also allow specific table-related PAP sprms like sprmPItap (0x6649, sgc=3)
        if (sgc != 1 && sgc != 2 && sprm.Code != 0x6649 && sprm.Code != 0x2416) return;

        switch (sprmCode)
        {
            // --- Word 97+ (16-bit) SPRM Opcodes ---
            case 0x00: pap.StyleId = (ushort)sprm.Operand; break; // sprmPIstd
            case 0x03: pap.KeepWithNext = sprm.Operand != 0; break; // sprmPFKeep
            case 0x04: pap.KeepTogether = sprm.Operand != 0; break; // sprmPFKeepFollow
            case 0x05: pap.PageBreakBefore = sprm.Operand != 0; break; // sprmPPageBreakBefore
            case 0x0B: pap.ListFormatId = (int)(short)sprm.Operand; break; // sprmPIlfo
            case 0x0A: pap.ListLevel = (byte)sprm.Operand; break; // sprmPIlvl
            case 0x0E: pap.IndentRight = (int)(short)sprm.Operand; break; // sprmPDxaRight
            case 0x0F: pap.IndentLeft = (int)(short)sprm.Operand; break; // sprmPDxaLeft
            case 0x11: pap.IndentFirstLine = (int)(short)sprm.Operand; break; // sprmPDxaLeft1
            case 0x12: // sprmPDyaLine — LSPD structure: low 16 bits = dyaLine (signed), bit 16 = fMultLinespace
                pap.LineSpacing = (int)(short)(sprm.Operand & 0xFFFF);
                pap.LineSpacingMultiple = (int)((sprm.Operand >> 16) & 1);
                break;
            case 0x13: pap.SpaceBefore = (int)(short)sprm.Operand; break; // sprmPDyaBefore
            case 0x14: pap.SpaceAfter = (int)(short)sprm.Operand; break; // sprmPDyaAfter
            case 0x16: pap.InTable = sprm.Operand != 0; break; // sprmPFInTable (0x2416 -> code 0x16)
            case 0x40: pap.OutlineLevel = (byte)sprm.Operand; break; // sprmPOutlineLvl
            case 0x49: pap.Itap = (int)(short)sprm.Operand; break; // sprmPItap (0x6649 -> code 0x49)
            case 0x61: pap.Justification = (byte)sprm.Operand; break; // sprmPJc
            // sprmPShd — paragraph shading (SHDOperand or Shd)
            case 0x0C:
                if (sprm.VariableOperand != null && sprm.VariableOperand.Length >= 4)
                {
                    try
                    {
                        ParseShdOperand(sprm.VariableOperand, out var fore, out var back, out var patternVal);
                        pap.Shading = new ShadingInfo
                        {
                            ForegroundColor = fore,
                            BackgroundColor = back,
                            PatternVal = patternVal
                        };
                    }
                    catch { /* ignore */ }
                }
                break;

            // --- Word 6 (8-bit) SPRM Opcodes (Fallbacks) ---
            case 0x02: pap.StyleId = (ushort)sprm.Operand; break;
            case 0x15: pap.LineSpacing = (int)sprm.Operand; break;
            // case 0x16: pap.SpaceBefore = (int)sprm.Operand; break; // Conflicts with Word 97 sprmPFInTable
            case 0x17: pap.SpaceAfter = (int)sprm.Operand; break;
        }
    }

    private void ApplyTapSprm(Sprm sprm, TapBase tap)
    {
        var sprmCode = sprm.Code & 0x03FF;
        var sgc = (sprm.Code >> 13) & 0x07;
        if (sgc != 3) return;
        switch (sprmCode)
        {
            // Table indent from left margin (sprmTDxaLeft)
            // Operand is a signed 16‑bit twip value.
            case 0x01:
                tap.IndentLeft = (int)(short)sprm.Operand;
                break;

            // Half of the inter‑cell gap (sprmTDxaGapHalf). The effective
            // cell spacing between two adjacent cells is typically 2 * GapHalf.
            // We also update CellSpacing when it has not been set by other
            // TAP sprms so table layout code has a single, easy source.
            case 0x02:
                tap.GapHalf = (int)(short)sprm.Operand;
                if (tap.CellSpacing == 0)
                {
                    tap.CellSpacing = tap.GapHalf * 2;
                }
                break;
            case 0x03: tap.CantSplit = sprm.Operand != 0; break; // sprmTFCantSplit
            case 0x04: tap.IsHeaderRow = sprm.Operand != 0; break; // sprmTHeader
            // sprmTTableBorders — table-wide borders (top/bottom/left/right/insideH/insideV).
            // The variable operand is an array of 6 BRC structures in the order:
            // top, left, bottom, right, insideH, insideV. Each BRC is 4 bytes in
            // the Word 97 binary format. We decode width/style/color into BorderInfo.
            case 0x05:
                if (sprm.VariableOperand != null && sprm.VariableOperand.Length >= 6 * 4)
                {
                    try
                    {
                        using var brcMs = new MemoryStream(sprm.VariableOperand);
                        using var brcReader = new BinaryReader(brcMs);
                        var borders = new BorderInfo[6];
                        for (int i = 0; i < 6; i++)
                        {
                            if (brcReader.BaseStream.Position + 4 > brcReader.BaseStream.Length)
                                break;

                            var brc = brcReader.ReadUInt32();
                            borders[i] = DecodeBrc(brc);
                        }

                        if (borders[0] != null) tap.BorderTop = borders[0];
                        if (borders[2] != null) tap.BorderBottom = borders[2];
                        if (borders[1] != null) tap.BorderLeft = borders[1];
                        if (borders[3] != null) tap.BorderRight = borders[3];
                        if (borders[4] != null) tap.BorderInsideH = borders[4];
                        if (borders[5] != null) tap.BorderInsideV = borders[5];
                    }
                    catch
                    {
                        // best-effort only
                    }
                }
                break;
            case 0x06: tap.Justification = (byte)sprm.Operand; break; // sprmTJc
            case 0x07: break;
            case 0x08: // sprmTDefTable - cell boundaries and TC (cell) descriptors
                if (sprm.VariableOperand != null && sprm.VariableOperand.Length > 0)
                {
                    try
                    {
                        using var defMs = new MemoryStream(sprm.VariableOperand);
                        using var defReader = new BinaryReader(defMs);
                        var cellCount = defReader.ReadByte();
                        if (cellCount > 0 && defMs.Length >= 1 + (cellCount + 1) * 2)
                        {
                            // Read cell boundary positions (in twips)
                            var boundaries = new short[cellCount + 1];
                            for (int i = 0; i <= cellCount; i++)
                                boundaries[i] = defReader.ReadInt16();
                            
                            // Calculate cell widths from boundary differences
                            tap.CellWidths = new int[cellCount];
                            for (int i = 0; i < cellCount; i++)
                                tap.CellWidths[i] = Math.Abs(boundaries[i + 1] - boundaries[i]);

                            // After rgdxaCenter comes rgTc (TC structures). The exact size of TC
                            // can vary between Word versions; for our purposes we only require
                            // the first 2 bytes (grfw bitfield), which contain the merge flags.
                            // We conservatively assume a fixed TC size and advance cautiously.
                            var merges = new CellMergeFlags[cellCount];
                            const int assumedTcSize = 20; // bytes per TC (>= 2, includes padding/other fields)

                            for (int i = 0; i < cellCount; i++)
                            {
                                if (defMs.Position + 2 > defMs.Length)
                                {
                                    break;
                                }

                                var grfw = defReader.ReadUInt16();
                                var flags = new CellMergeFlags
                                {
                                    // Bit meanings follow the TC.grfw definition in MS‑DOC:
                                    // 0x0001 = fFirstMerged, 0x0002 = fMerged,
                                    // 0x0004 = fFirstVertMerge, 0x0008 = fVertMerge.
                                    HorizFirst = (grfw & 0x0001) != 0,
                                    HorizMerged = (grfw & 0x0002) != 0,
                                    VertFirst = (grfw & 0x0004) != 0,
                                    VertMerged = (grfw & 0x0008) != 0
                                };
                                merges[i] = flags;

                                // Skip the remainder of the TC structure, if present.
                                if (assumedTcSize > 2 && defMs.Position + (assumedTcSize - 2) <= defMs.Length)
                                {
                                    defMs.Seek(assumedTcSize - 2, SeekOrigin.Current);
                                }
                                else if (defMs.Position > defMs.Length)
                                {
                                    break;
                                }
                            }

                            tap.CellMerges = merges;
                        }
                    }
                    catch
                    {
                        // Ignore parse errors; table layout will fall back to widths only.
                    }
                }
                break;
            case 0x09: break; // sprmTSetBrc (cell borders) — not yet mapped per-cell in this version
            case 0x0A: break;
            case 0x0B: break;
            case 0x0C: break;
            // sprmTShd — table shading. Operand can be SHDOperand (cb=10 + Shd 10 bytes) or short form.
            case 0x0D:
                if (sprm.VariableOperand != null && sprm.VariableOperand.Length >= 4)
                {
                    try
                    {
                        ParseShdOperand(sprm.VariableOperand, out var fore, out var back, out var patternVal);
                        tap.Shading ??= new ShadingInfo();
                        tap.Shading.ForegroundColor = fore;
                        tap.Shading.BackgroundColor = back;
                        if (patternVal != null)
                            tap.Shading.PatternVal = patternVal;
                    }
                    catch
                    {
                        // ignore shading parse errors
                    }
                }
                break;
            case 0x0E: break;
            case 0x0F: break;
            case 0x10: tap.RowHeight = (int)sprm.Operand; break;
            case 0x11: tap.HeightIsExact = sprm.Operand != 0; break;
            case 0x12: break;
            case 0x13: tap.CellSpacing = (int)(short)sprm.Operand; break;
            case 0x14: tap.TableWidth = (int)(short)sprm.Operand; break;
            case 0x15: break;
            case 0x16: break;
            case 0x17: break;
            case 0x18: break;
            case 0x19: break;
            case 0x1A: break;
            case 0x1B: break;
            case 0x1C: break;
            case 0x1D: break;
            case 0x1E: break;
            case 0x1F: break;
        }
    }

    /// <summary>
    /// Parses SHD/SHDOperand: full Shd (cvFore 4, cvBack 4, ipat 2) or legacy icoFore/icoBack (2+2).
    /// COLORREF is 0x00BBGGRR; we output RGB as int for ColorHelper.
    /// </summary>
    private static void ParseShdOperand(byte[] data, out int foreColor, out int backColor, out string? patternVal)
    {
        foreColor = 0;
        backColor = 0;
        patternVal = null;
        int offset = 0;
        if (data.Length == 11 && data[0] == 10)
            offset = 1; // SHDOperand: cb=10, then Shd
        if (data.Length >= offset + 10)
        {
            var cvFore = BitConverter.ToUInt32(data, offset);
            var cvBack = BitConverter.ToUInt32(data, offset + 4);
            var ipat = BitConverter.ToUInt16(data, offset + 8);
            // Store as int; ColorHelper.ColorToHex expects COLORREF (0x00BBGGRR) for values > 16
            foreColor = (int)cvFore;
            backColor = (int)cvBack;
            patternVal = IpatToShdVal(ipat);
            return;
        }
        if (data.Length >= offset + 4)
        {
            foreColor = BitConverter.ToUInt16(data, offset);
            backColor = BitConverter.ToUInt16(data, offset + 2);
            patternVal = "clear";
        }
    }

    /// <summary>Maps MS-DOC Ipat to OOXML w:shd val per [MS-DOC] 2.9.121.</summary>
    private static string? IpatToShdVal(ushort ipat)
    {
        if (ipat == 0xFFFF) return "nil";
        return ipat switch
        {
            0 => "clear",
            1 => "solid",
            2 => "pct5",
            3 => "pct10",
            4 => "pct20",
            5 => "pct25",
            6 => "pct30",
            7 => "pct40",
            8 => "pct50",
            9 => "pct60",
            0x0A => "pct70",
            0x0B => "pct75",
            0x0C => "pct80",
            0x0D => "pct90",
            0x0E => "horzStripe",
            0x0F => "vertStripe",
            0x10 => "reverseDiagStripe",
            0x11 => "diagStripe",
            0x12 => "horzCross",
            0x13 => "diagCross",
            0x14 => "thinHorzStripe",
            0x15 => "thinVertStripe",
            0x16 => "thinReverseDiagStripe",
            0x17 => "thinDiagStripe",
            0x18 => "thinHorzCross",
            0x19 => "thinDiagCross",
            0x25 => "pct12",
            0x26 => "pct15",
            0x2B => "pct35",
            0x2C => "pct37",
            0x2E => "pct45",
            0x31 => "pct55",
            0x33 => "pct62",
            0x34 => "pct65",
            0x39 => "pct85",
            0x3C => "pct95",
            _ => "clear"
        };
    }

    /// <summary>
    /// Decodes a Word binary BRC80 value into a high-level BorderInfo.
    /// Per MS-DOC section 2.9.16 (Brc80):
    ///  - bits 0-7   (8 bits) = dptLineWidth (width in 1/8 pt)
    ///  - bits 8-15  (8 bits) = brcType (border style index)
    ///  - bits 16-23 (8 bits) = ico (color index)
    ///  - bits 24-28 (5 bits) = dptSpace (spacing in pt)
    ///  - bit  29    = fShadow
    ///  - bit  30    = fFrame
    /// </summary>
    private static BorderInfo DecodeBrc(uint brc)
    {
        var dptLineWidth = (int)(brc & 0xFF);         // bits 0-7
        var brcType      = (int)((brc >> 8) & 0xFF);  // bits 8-15
        var ico          = (int)((brc >> 16) & 0xFF);  // bits 16-23
        var dptSpace     = (int)((brc >> 24) & 0x1F);  // bits 24-28

        var style = brcType switch
        {
            0 => BorderStyle.None,
            1 => BorderStyle.Single,
            2 => BorderStyle.Thick,
            3 => BorderStyle.Double,
            5 => BorderStyle.Dotted,         // hairline
            6 => BorderStyle.Dashed,
            7 => BorderStyle.DotDash,
            8 => BorderStyle.DotDotDash,
            9 => BorderStyle.Triple,
            10 => BorderStyle.ThinThickSmallGap,
            11 => BorderStyle.ThickThinSmallGap,
            12 => BorderStyle.ThinThickThinSmallGap,
            13 => BorderStyle.ThinThickMediumGap,
            14 => BorderStyle.ThickThinMediumGap,
            15 => BorderStyle.ThinThickThinMediumGap,
            16 => BorderStyle.ThinThickLargeGap,
            17 => BorderStyle.ThickThinLargeGap,
            18 => BorderStyle.ThinThickThinLargeGap,
            19 => BorderStyle.Wave,
            _ => BorderStyle.Single
        };

        // dptLineWidth is in 1/8 pt; OOXML w:sz is in 1/8 pt, so use directly
        var widthEighthPt = dptLineWidth;

        return new BorderInfo
        {
            Style = style,
            Width = widthEighthPt,
            Color = ico,
            Space = dptSpace
        };
    }

    private void ApplySepSprm(Sprm sprm, SepBase sep)
    {
        // bits 13-15 = sgc. For Section, sgc = 4.
        var sgc = (sprm.Code >> 13) & 0x07;
        if (sgc != 4) return;

        var sprmCode = sprm.Code & 0x01FF;
        switch (sprmCode)
        {
            case 0x00: sep.BreakCode = (byte)sprm.Operand; break; // sprmSBkc
            case 0x01: sep.TitlePage = sprm.Operand != 0; break; // sprmSFTitlePage
            case 0x02: sep.ColumnCount = (short)(sprm.Operand + 1); break; // sprmSCColumns
            case 0x03: sep.ColumnSpacing = (int)(short)sprm.Operand; break; // sprmSDxaColumns
            case 0x0F: sep.PageWidth = (int)(short)sprm.Operand; break; // sprmSDxaPage
            case 0x10: sep.PageHeight = (int)(short)sprm.Operand; break; // sprmSDyaPage
            case 0x11: sep.MarginLeft = (int)(short)sprm.Operand; break; // sprmSDxaLeft
            case 0x12: sep.MarginRight = (int)(short)sprm.Operand; break; // sprmSDxaRight
            case 0x13: sep.MarginTop = (int)(short)sprm.Operand; break; // sprmSDyaTop
            case 0x14: sep.MarginBottom = (int)(short)sprm.Operand; break; // sprmSDyaBottom
            case 0x15: sep.MarginHeader = (int)(short)sprm.Operand; break; // sprmSDzaHdrTop
            case 0x16: sep.MarginFooter = (int)(short)sprm.Operand; break; // sprmSDzaHdrBottom
            case 0x17: sep.Gutter = (int)(short)sprm.Operand; break; // sprmSDxaGutter
            case 0x2A: sep.VerticalAlignment = (byte)sprm.Operand; break; // sprmSVjc
        }
    }

    private class Sprm
    {
        public ushort Code { get; set; }
        public int OperandSize { get; set; }
        public uint Operand { get; set; }
        public byte[]? VariableOperand { get; set; }
    }
}

public class ChpBase
{
    public short FontIndex { get; set; } = -1;
    public byte FontSize { get; set; } = 24;
    public byte FontSizeCs { get; set; } = 24;
    public bool IsBold { get; set; }
    public bool IsBoldCs { get; set; }
    public bool IsItalic { get; set; }
    public bool IsItalicCs { get; set; }
    public bool IsUnderline { get; set; }
    public byte Underline { get; set; }
    public bool IsStrikeThrough { get; set; }
    public bool IsSmallCaps { get; set; }
    public bool IsAllCaps { get; set; }
    public bool IsHidden { get; set; }
    public bool IsSuperscript { get; set; }
    public bool IsSubscript { get; set; }
    public byte Color { get; set; }
    public short FontIndexCs { get; set; } = -1;
    public int CharacterSpacingAdjustment { get; set; }
    public int Language { get; set; }
    public int LanguageId { get; set; }
    public bool IsDoubleStrikeThrough { get; set; }
    public int DxaOffset { get; set; }
    // Phase 3 additions
    public bool IsOutline { get; set; }
    public int Kerning { get; set; }
    public int Position { get; set; }
    /// <summary>File character offset in Data stream for picture (sprmCPicLocation).</summary>
    public uint FcPic { get; set; }
    public int Scale { get; set; } = 100;
    public byte HighlightColor { get; set; }
    public bool IsShadow { get; set; }
    public bool IsEmboss { get; set; }
    public bool IsImprint { get; set; }
    public uint RgbColor { get; set; }
    public bool HasRgbColor { get; set; }
    
    // Track Changes
    public bool IsDeleted { get; set; }
    public bool IsInserted { get; set; }
    public ushort AuthorIndexDel { get; set; }
    public ushort AuthorIndexIns { get; set; }
    public uint DateDel { get; set; }
    public uint DateIns { get; set; }
}

public class PapBase
{
    public ushort StyleId { get; set; }
    public ushort Istd { get; set; }
    public byte Justification { get; set; }
    public bool KeepWithNext { get; set; }
    public bool KeepTogether { get; set; }
    public bool PageBreakBefore { get; set; }
    public int IndentLeft { get; set; }
    public int IndentRight { get; set; }
    public int IndentFirstLine { get; set; }
    public int LineSpacing { get; set; } = 240;
    public int LineSpacingMultiple { get; set; }
    public int SpaceBefore { get; set; }
    public int SpaceAfter { get; set; }
    // Phase 3 additions
    public byte OutlineLevel { get; set; } = 9; // 9 = body text
    public int NestIndent { get; set; }
    public int ListFormatId { get; set; }
    public byte ListLevel { get; set; }
    public int ListFormatOverrideId { get; set; }
    /// <summary>Paragraph-level shading (background/pattern) when sprmPShd is present.</summary>
    public ShadingInfo? Shading { get; set; }
    // Nested table info
    public bool InTable { get; set; }
    public int Itap { get; set; }
    // Associated table properties (TAP) decoded from the same GRPPRL, when present.
    public TapBase? Tap { get; set; }
}

public class TapBase
{
    public int RowHeight { get; set; }
    public bool HeightIsExact { get; set; }
    // Phase 3 additions
    /// <summary>
    /// Table justification (left/center/right) as stored in TAP.
    /// </summary>
    public byte Justification { get; set; }
    /// <summary>
    /// True when this row is marked as a header row that should repeat on each page.
    /// </summary>
    public bool IsHeaderRow { get; set; }
    /// <summary>
    /// Cell spacing in twips (total distance between cell borders).
    /// </summary>
    public int CellSpacing { get; set; }
    /// <summary>
    /// Preferred table width in twips, if specified.
    /// </summary>
    public int TableWidth { get; set; }
    /// <summary>
    /// Absolute left indent of the table from the page/column margin, in twips.
    /// </summary>
    public int IndentLeft { get; set; }
    /// <summary>
    /// Half of the inter‑cell gap (TDxaGapHalf); when present, the effective
    /// cell spacing is typically 2 * GapHalf. We keep both GapHalf and the
    /// derived CellSpacing so callers can choose the most appropriate value.
    /// </summary>
    public int GapHalf { get; set; }
    /// <summary>
    /// Per‑cell widths in twips, derived from the TAP boundary positions.
    /// </summary>
    public int[]? CellWidths { get; set; }
    /// <summary>
    /// When true, the row must not be split across pages (cantSplit).
    /// </summary>
    public bool CantSplit { get; set; }

    /// <summary>
    /// Per‑cell merge flags decoded from the TC structures that follow TDefTable.
    /// This captures both horizontal (grid) and vertical merge intentions as stored
    /// in the binary document.
    /// </summary>
    public CellMergeFlags[]? CellMerges { get; set; }

    /// <summary>
    /// Table-level borders derived from sprmTTableBorders, mapped into the same
    /// shape as the high-level TableProperties so that table writers can reuse
    /// a single mapping path.
    /// </summary>
    public BorderInfo? BorderTop { get; set; }
    public BorderInfo? BorderBottom { get; set; }
    public BorderInfo? BorderLeft { get; set; }
    public BorderInfo? BorderRight { get; set; }
    public BorderInfo? BorderInsideH { get; set; }
    public BorderInfo? BorderInsideV { get; set; }

    /// <summary>
    /// Table-level shading (background) for the whole table, when present.
    /// </summary>
    public ShadingInfo? Shading { get; set; }
}

/// <summary>
/// Merge flags for a single table cell, as decoded from the TC.grfw bitfield
/// in the TAP/row formatting. The exact semantics follow the MS‑DOC TC structure:
/// fFirstMerged/fMerged for horizontal merges; fFirstVertMerge/fVertMerge for
/// vertical merges. We expose them in a neutral form here so higher‑level
/// table reconstruction code can decide how to map them into RowSpan/ColumnSpan.
/// </summary>
public class CellMergeFlags
{
    /// <summary>True if this cell is the first cell in a horizontal merge sequence.</summary>
    public bool HorizFirst { get; set; }

    /// <summary>True if this cell is horizontally merged into a previous cell.</summary>
    public bool HorizMerged { get; set; }

    /// <summary>True if this cell is the first cell in a vertical merge sequence.</summary>
    public bool VertFirst { get; set; }

    /// <summary>True if this cell is vertically merged into a cell above.</summary>
    public bool VertMerged { get; set; }
}

public class SepBase
{
    public byte BreakCode { get; set; } // SBkc (0=cont, 1=col, 2=page, 3=even, 4=odd)
    public bool TitlePage { get; set; } // SFTitlePage
    public short ColumnCount { get; set; } = 1; // SCColumns
    public int ColumnSpacing { get; set; } // SDxaColumns
    public int PageWidth { get; set; } = 11906; // 21cm (A4)
    public int PageHeight { get; set; } = 16838; // 29.7cm (A4)
    public int MarginLeft { get; set; } = 1440; // 1"
    public int MarginRight { get; set; } = 1440;
    public int MarginTop { get; set; } = 1440;
    public int MarginBottom { get; set; } = 1440;
    public int MarginHeader { get; set; } = 720;
    public int MarginFooter { get; set; } = 720;
    public int Gutter { get; set; }
    public byte VerticalAlignment { get; set; } // SVjc (0=top, 1=center, 2=justified, 3=bottom)
}
