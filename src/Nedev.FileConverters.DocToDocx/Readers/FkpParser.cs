using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// FKP (Formatted Disk Page) parser for Word 97-2003 binary format.
/// FKPs store character (CHP) and paragraph (PAP) formatting properties.
/// </summary>
public class FkpParser
{
    private readonly BinaryReader _wordDocReader;
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly SprmParser _sprmParser;
    private readonly TextReader _textReader;

    private readonly Dictionary<uint, ChpFkp> _chpFkpCache = new();
    private readonly Dictionary<uint, PapFkp> _papFkpCache = new();

    // cache of FC↔CP segments gleaned from parsed FKP pages; used by
    // CpToFc() to locate byte offsets for given CP positions when no piece
    // table is available (simple documents with footnotes).
    private readonly List<(int cpStart, int cpEnd, int fcStart, int fcEnd)> _fcCpSegments = new();

    public FkpParser(BinaryReader wordDocReader, BinaryReader tableReader, FibReader fib, TextReader textReader)
    {
        _wordDocReader = wordDocReader;
        _tableReader = tableReader;
        _fib = fib;
        _textReader = textReader;
        _sprmParser = new SprmParser(wordDocReader, 0);
    }

    #region CHP (Character Properties)

    public Dictionary<int, ChpBase> ReadChpProperties()
    {
            // clear any previous FC↔CP segments when we re-read the table
            _fcCpSegments.Clear();

        var chpMap = new Dictionary<int, ChpBase>();
        _tableReader.BaseStream.Seek(_fib.FcPlcfBteChpx, SeekOrigin.Begin);
        
        // BTE PLC structure: CP array (n+1 entries) + PN array (n entries).
        // Both entries are 4 bytes here, so total size = 4 + n*8.
        var lcb = (int)_fib.LcbPlcfBteChpx;
        var numPcd = (lcb - 4) / 8;
        if (numPcd <= 0) return chpMap;
        
        // Read CP array: numPcd + 1 entries
        var cpArray = new int[numPcd + 1];
        int cpRead = 0;
        for (int i = 0; i <= numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            cpArray[i] = _tableReader.ReadInt32();
            cpRead++;
        }
        
        // Read PCD array: numPcd entries (PNs)
        var pnArray = new uint[numPcd];
        int pnRead = 0;
        for (int i = 0; i < numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            pnArray[i] = _tableReader.ReadUInt32();
            pnRead++;
        }
        
        // only iterate over entries we actually read; avoids huge loops when
        // header length was corrupted
        var loopCount = Math.Min(cpRead - 1, pnRead);
        var storyCpLimit = GetDocumentCpLimit();
        for (int i = 0; i < loopCount; i++)
        {
            var fkp = GetChpFkp(pnArray[i]);
            if (fkp == null) continue;
            
            foreach (var entry in fkp.Entries)
            {
                // ParseChpFkp already normalized the FKP FC bounds into CP bounds.
                int startCp = entry.StartCpOffset;
                int endCp = entry.EndCpOffset;
                
                int finalStart = Math.Max(0, startCp);
                int finalEnd = Math.Min(storyCpLimit, endCp);
                
                for (int cp = finalStart; cp < finalEnd; cp++) chpMap[cp] = entry.Properties;
            }
        }
        return chpMap;
    }

    private ChpFkp? GetChpFkp(uint pn)
    {
        if (_chpFkpCache.TryGetValue(pn, out var cached)) return cached;
        var fkp = LoadChpFkp(pn);
        if (fkp != null) _chpFkpCache[pn] = fkp;
        return fkp;
    }

    private ChpFkp? LoadChpFkp(uint pn)
    {
        var offset = pn * WordConsts.FKP_PAGE_SIZE;
        if (offset + WordConsts.FKP_PAGE_SIZE > _wordDocReader.BaseStream.Length) return null;
        _wordDocReader.BaseStream.Seek(offset, SeekOrigin.Begin);
        return ParseChpFkp(_wordDocReader.ReadBytes(WordConsts.FKP_PAGE_SIZE));
    }

    private ChpFkp ParseChpFkp(byte[] data)
    {
        var fkp = new ChpFkp();
        
        // Safety check: data must be FKP_PAGE_SIZE bytes
        if (data.Length < WordConsts.FKP_PAGE_SIZE) return fkp;
        
        // Per MS-DOC spec: crun is the LAST byte of the 512-byte FKP page
        var crun = data[WordConsts.FKP_PAGE_SIZE - 1];
        if (crun == 0 || crun > 101) return fkp; // max crun for CHPX FKP is 101

        // FKP layout:
        //   rgfc[0..crun] : (crun+1) x 4-byte FC/CP values at offset 0
        //   rgb[0..crun-1] : crun x 1-byte offsets (word offsets into this page)
        //   ... property data ...
        //   crun : 1 byte at offset 511
        
        var rgfcSize = (crun + 1) * 4;
        if (rgfcSize + crun > data.Length) return fkp;

        // Read FC/CP array (crun+1 entries, starting at offset 0)
        var fcArray = new int[crun + 1];
        for (int i = 0; i <= crun; i++)
        {
            fcArray[i] = BitConverter.ToInt32(data, i * 4);
        }

        // Read property offset bytes (crun entries, starting after FC array)
        var rgbBase = rgfcSize;

        for (int i = 0; i < crun; i++)
        {
            if (rgbBase + i >= data.Length) break;
            
            var propOffset = data[rgbBase + i];
            var dataOffset = propOffset * 2;
            var cb = 0;
            
            var chp = new ChpBase();
            var grpprl = Array.Empty<byte>();

            if (propOffset != 0)
            {
                // propOffset is a word offset (multiply by 2 to get byte offset)
                if (dataOffset < WordConsts.FKP_PAGE_SIZE && dataOffset < data.Length)
                {
                    cb = data[dataOffset];
                    if (cb > 0 && dataOffset + 1 + cb <= data.Length)
                    {
                        grpprl = new byte[cb];
                        Array.Copy(data, dataOffset + 1, grpprl, 0, cb);
                        _sprmParser.ApplyToChp(grpprl, chp);
                    }
                }
            }

            var startFc = fcArray[i];
            var endFc = fcArray[i + 1];
            var cpStart = FcToCp(startFc);
            var cpEnd = FcToCp(endFc);
            fkp.Entries.Add(new ChpFkpEntry
            {
                StartCpOffset = cpStart,
                EndCpOffset = cpEnd,
                StartFcOffset = startFc,
                EndFcOffset = endFc,
                RawGrpprl = grpprl,
                Properties = chp
            });
            // remember segment for CP→FC lookups later
            _fcCpSegments.Add((cpStart, cpEnd, startFc, endFc));
        }
        return fkp;
    }

    #endregion

    #region PAP (Paragraph Properties)

    public Dictionary<int, PapBase> ReadPapProperties()
    {
        var papMap = new Dictionary<int, PapBase>();
        if (_fib.FcPlcfBtePapx == 0 || _fib.LcbPlcfBtePapx < 16) return papMap;

        _tableReader.BaseStream.Seek(_fib.FcPlcfBtePapx, SeekOrigin.Begin);
        
        // BTE PLC structure: CP array (n+1 entries) + PN array (n entries).
        // Both entries are 4 bytes here, so total size = 4 + n*8.
        var lcb = (int)_fib.LcbPlcfBtePapx;
        var numPcd = (lcb - 4) / 8;
        if (numPcd <= 0) return papMap;
        
        // Read CP array: numPcd + 1 entries
        var cpArray = new int[numPcd + 1];
        int cpRead = 0;
        for (int i = 0; i <= numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            cpArray[i] = _tableReader.ReadInt32();
            cpRead++;
        }
        
        // Read PCD array: numPcd entries (PNs)
        var pnArray = new uint[numPcd];
        int pnRead = 0;
        for (int i = 0; i < numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            pnArray[i] = _tableReader.ReadUInt32();
            pnRead++;
        }
        
        var loopCount = Math.Min(cpRead - 1, pnRead);
        var storyCpLimit = GetDocumentCpLimit();
        for (int i = 0; i < loopCount; i++)
        {
            var fkp = GetPapFkp(pnArray[i]);
            if (fkp == null) continue;
            
            foreach (var entry in fkp.Entries)
            {
                // ParsePapFkp already converted the FKP range from FC into CP.
                int startCp = entry.StartCpOffset;
                int endCp = entry.EndCpOffset;
                
                int finalStart = Math.Max(0, startCp);
                int finalEnd = Math.Min(storyCpLimit, endCp);
                
                for (int cp = finalStart; cp < finalEnd; cp++) papMap[cp] = entry.Properties;
            }
        }
        return papMap;
    }

    private PapFkp? GetPapFkp(uint pn)
    {
        if (_papFkpCache.TryGetValue(pn, out var cached)) return cached;
        var fkp = LoadPapFkp(pn);
        if (fkp != null) _papFkpCache[pn] = fkp;
        return fkp;
    }

    private PapFkp? LoadPapFkp(uint pn)
    {
        try
        {
            var offset = pn * WordConsts.FKP_PAGE_SIZE;
            if (offset + WordConsts.FKP_PAGE_SIZE > _wordDocReader.BaseStream.Length) return null;
            _wordDocReader.BaseStream.Seek(offset, SeekOrigin.Begin);
            return ParsePapFkp(_wordDocReader.ReadBytes(WordConsts.FKP_PAGE_SIZE));
        }
        catch
        {
            return null;
        }
    }

    private PapFkp ParsePapFkp(byte[] data)
    {
        var fkp = new PapFkp();
        
        // Safety check: data must be FKP_PAGE_SIZE bytes
        if (data.Length < WordConsts.FKP_PAGE_SIZE) return fkp;
        
        // Per MS-DOC spec: cpara (crun) is the LAST byte of the 512-byte FKP page
        var crun = data[WordConsts.FKP_PAGE_SIZE - 1];
        if (crun == 0 || crun > 101) return fkp;

        // FKP layout:
        //   rgfc[0..crun] : (crun+1) x 4-byte FC/CP values at offset 0
        //   rgbx[0..crun-1] : crun x 13-byte BX entries (for PAPX)
        //     Each BX: 1 byte (offset), 12 bytes (PHE descriptor)
        //   ... property data ...
        //   cpara : 1 byte at offset 511
        
        var rgfcSize = (crun + 1) * 4;
        if (rgfcSize > data.Length) return fkp;

        // Read FC/CP array (crun+1 entries, starting at offset 0)
        var fcArray = new int[crun + 1];
        for (int i = 0; i <= crun; i++)
        {
            fcArray[i] = BitConverter.ToInt32(data, i * 4);
        }

        // BX entries start right after the FC array
        // Each BX is 13 bytes for PAPX FKP (1 byte offset + 12 bytes PHE)
        var bxBase = rgfcSize;
        var bxSize = 13; // PAPX BX size

        for (int i = 0; i < crun; i++)
        {
            var bxOffset = bxBase + i * bxSize;
            if (bxOffset >= data.Length) break;
            
            var props = new PapBase();
            var bx = data[bxOffset]; // First byte of BX is the word offset
            var dataOffset = bx * 2;
            if (dataOffset < WordConsts.FKP_PAGE_SIZE && dataOffset < data.Length)
            {
                // PAPX size is stored as a word count. Word stores either
                // (cb * 2) - 1 bytes directly or, for aligned edge cases, a
                // zero byte followed by the true word count.
                var cb = data[dataOffset] * 2;
                if (cb == 0 && dataOffset + 1 < data.Length)
                {
                    dataOffset++;
                    cb = data[dataOffset] * 2;
                }
                else if (cb > 0)
                {
                    cb--;
                }

                if (cb >= 2 && dataOffset + 1 + cb <= data.Length)
                {
                    props.Istd = BitConverter.ToUInt16(data, dataOffset + 1);
                    var grpprlLength = cb - 2;
                    if (grpprlLength > 0 && dataOffset + 3 + grpprlLength <= data.Length)
                    {
                        var grpprl = new byte[grpprlLength];
                        Array.Copy(data, dataOffset + 3, grpprl, 0, grpprlLength);
                        // Decode paragraph and table (TAP) properties from the same GRPPRL.
                        _sprmParser.ApplyToPap(grpprl, props);
                        var tap = new TapBase();
                        _sprmParser.ApplyToTap(grpprl, tap);
                        props.Tap = tap;
                    }
                }
            }

            if (i + 1 < fcArray.Length)
            {
                var startFc = fcArray[i];
                var endFc = fcArray[i + 1];
                var cpStart = FcToCp(startFc);
                var cpEnd = FcToCp(endFc);

                fkp.Entries.Add(new PapFkpEntry
                {
                    StartCpOffset = cpStart,
                    EndCpOffset = cpEnd,
                    Properties = props
                });
            }
        }
        return fkp;
    }

    #endregion

    #region FC to CP conversion

    private int FcToCp(int fc)
    {
        if (_textReader == null || _textReader.Pieces.Count == 0)
        {
            // Fallback for simple documents (no Piece Table). These can still
            // use double-byte encodings such as GBK, so FC is not necessarily a
            // one-to-one match with decoded character positions.
            if (fc < _fib.FcMin) return 0;

            if (_textReader != null)
                return _textReader.MapSimpleFcToCp(fc);

            return (int)(fc - _fib.FcMin);
        }

        foreach (var piece in _textReader.Pieces)
        {
            var pieceFc = (int)piece.FileOffset;
            var bytesPerChar = piece.IsUnicode ? 2 : 1;
            var pieceLengthBytes = piece.CharCount * bytesPerChar;
            
            if (fc >= pieceFc && fc < pieceFc + pieceLengthBytes)
            {
                var offsetInPiece = fc - pieceFc;
                return piece.CpStart + (offsetInPiece / bytesPerChar);
            }

            // Some structures expose FC values in the raw PCD encoding rather than
            // the decoded byte offset used by TextReader.FileOffset. For compressed
            // pieces RawFcMasked advances by 2 per character even though the actual
            // byte stream advances by 1. Try this representation as a fallback.
            var rawPieceFc = (int)piece.RawFcMasked;
            var rawUnitsPerChar = 2;
            var rawPieceLength = piece.CharCount * rawUnitsPerChar;
            if (fc >= rawPieceFc && fc < rawPieceFc + rawPieceLength)
            {
                var rawOffsetInPiece = fc - rawPieceFc;
                return piece.CpStart + (rawOffsetInPiece / rawUnitsPerChar);
            }
        }
        
        // If not found, check if it's the exact end of the last piece
        var lastPiece = _textReader.Pieces.LastOrDefault();
        if (lastPiece != null)
        {
            var pieceFc = (int)lastPiece.FileOffset;
            var bytesPerChar = lastPiece.IsUnicode ? 2 : 1;
            var pieceLengthBytes = lastPiece.CharCount * bytesPerChar;
            
            if (fc == pieceFc + pieceLengthBytes)
            {
                return lastPiece.CpEnd;
            }

            var rawPieceFc = (int)lastPiece.RawFcMasked;
            var rawPieceLength = lastPiece.CharCount * 2;
            if (fc == rawPieceFc + rawPieceLength)
            {
                return lastPiece.CpEnd;
            }
        }

        // Fallback if no matching piece is found: clamp into valid CP range
        Logger.Warning($"FKP: Unmatched FC value {fc}, clamping to document range.");
        return Math.Clamp(fc, 0, _fib.CcpText);
    }

    /// <summary>
    /// Reverse mapping of <see cref="FcToCp"/>.  Uses the FC segments cached
    /// while parsing FKP pages to estimate a corresponding FC offset for a
    /// given CP position.  Returns null if nothing matches.
    /// </summary>
    public int? CpToFc(int cp)
    {
        foreach (var seg in _fcCpSegments)
        {
            if (cp >= seg.cpStart && cp < seg.cpEnd)
            {
                int cpLen = seg.cpEnd - seg.cpStart;
                int fcLen = seg.fcEnd - seg.fcStart;
                if (cpLen <= 0) return seg.fcStart;
                int bpc = Math.Max(1, fcLen / cpLen);
                return seg.fcStart + (cp - seg.cpStart) * bpc;
            }
        }
        return null;
    }

    private int GetDocumentCpLimit()
    {
        long total = (long)_fib.CcpText
            + _fib.CcpFtn
            + _fib.CcpHdd
            + _fib.CcpAtn
            + _fib.CcpEdn
            + _fib.CcpTxbx
            + _fib.CcpHdrTxbx;

        return total <= 0 ? _fib.CcpText : (int)Math.Min(int.MaxValue, total);
    }

    #endregion

    #region Convenience Methods

    public ChpBase? GetChpAtCp(int cp) => ReadChpProperties().TryGetValue(cp, out var chp) ? chp : null;
    public PapBase? GetPapAtCp(int cp) => ReadPapProperties().TryGetValue(cp, out var pap) ? pap : null;

    public RunProperties ConvertToRunProperties(ChpBase chp, StyleSheet styles)
    {
        var props = new RunProperties
        {
            FontIndex = chp.FontIndex,
            FontSize = chp.FontSize,
            FontSizeCs = chp.FontSizeCs,
            IsBold = chp.IsBold,
            IsBoldCs = chp.IsBoldCs,
            IsItalic = chp.IsItalic,
            IsItalicCs = chp.IsItalicCs,
            IsUnderline = chp.Underline != 0,
            UnderlineType = (UnderlineType)chp.Underline,
            IsStrikeThrough = chp.IsStrikeThrough,
            IsDoubleStrikeThrough = chp.IsDoubleStrikeThrough,
            IsSmallCaps = chp.IsSmallCaps,
            IsAllCaps = chp.IsAllCaps,
            IsHidden = chp.IsHidden,
            IsSuperscript = chp.IsSuperscript,
            IsSubscript = chp.IsSubscript,
            Color = chp.Color,
            CharacterSpacingAdjustment = chp.DxaOffset,
            Language = chp.LanguageId,
            // Phase 3 additions
            HighlightColor = chp.HighlightColor,
            RgbColor = chp.RgbColor,
            HasRgbColor = chp.HasRgbColor,
            IsOutline = chp.IsOutline,
            IsShadow = chp.IsShadow,
            IsEmboss = chp.IsEmboss,
            IsImprint = chp.IsImprint,
            Border = chp.Border,
            Kerning = chp.Kerning,
            Position = chp.Position,
            CharacterScale = chp.Scale,
            EastAsianLayoutType = chp.EastAsianLayoutType,
            IsEastAsianVertical = chp.IsEastAsianVertical,
            IsEastAsianVerticalCompress = chp.IsEastAsianVerticalCompress,
            
            // Track Changes
            IsDeleted = chp.IsDeleted,
            IsInserted = chp.IsInserted,
            AuthorIndexDel = chp.AuthorIndexDel,
            AuthorIndexIns = chp.AuthorIndexIns,
            DateDel = chp.DateDel,
            DateIns = chp.DateIns
        };
        if (chp.FontIndex >= 0 && chp.FontIndex < styles.Fonts.Count)
            props.FontName = styles.Fonts[chp.FontIndex].Name;
        return props;
    }

    public ParagraphProperties ConvertToParagraphProperties(PapBase pap, StyleSheet styles)
    {
        var styleIndex = pap.StyleId != 0 ? pap.StyleId : pap.Istd;
        return new ParagraphProperties
        {
            StyleIndex = styleIndex,
            Alignment = (ParagraphAlignment)pap.Justification,
            IndentLeft = pap.IndentLeft,
            IndentRight = pap.IndentRight,
            IndentFirstLine = pap.IndentFirstLine,
            SpaceBefore = pap.SpaceBefore,
            SpaceAfter = pap.SpaceAfter,
            LineSpacing = pap.LineSpacing,
            LineSpacingMultiple = pap.LineSpacingMultiple,
            KeepWithNext = pap.KeepWithNext,
            KeepTogether = pap.KeepTogether,
            PageBreakBefore = pap.PageBreakBefore,
            ListFormatId = pap.ListFormatId,
            ListLevel = pap.ListLevel,
            OutlineLevel = pap.OutlineLevel,
            Shading = pap.Shading
        };
    }

    #endregion
}

public class ChpFkp
{
    public List<ChpFkpEntry> Entries { get; set; } = new();
}

public class ChpFkpEntry
{
    public int StartCpOffset { get; set; }
    public int EndCpOffset { get; set; }

    // keep the original file-character (FC) offsets so that callers can map
    // a CP position back to a byte offset in the WordDocument stream.  this
    // is the inverse of the existing FcToCp helper, which we must invert for
    // simple documents where the piece table is missing.
    public int StartFcOffset { get; set; }
    public int EndFcOffset { get; set; }

    public byte[] RawGrpprl { get; set; } = Array.Empty<byte>();
    public ChpBase Properties { get; set; } = new();
}

public class PapFkp
{
    public List<PapFkpEntry> Entries { get; set; } = new();
}

public class PapFkpEntry
{
    public int StartCpOffset { get; set; }
    public int EndCpOffset { get; set; }
    public PapBase Properties { get; set; } = new();
}
