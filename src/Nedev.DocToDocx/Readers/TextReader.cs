using System.Linq;
using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Reads text content from Word 97-2003 binary documents.
/// Handles the CLX structure and Piece Table for both simple and complex documents.
/// 
/// Text in a .doc file is stored as a sequence of character positions (CPs).
/// For complex documents, the CLX structure in the Table stream contains
/// a Piece Table that maps CP ranges to file offsets in the WordDocument stream.
/// </summary>
public class TextReader
{
    private readonly BinaryReader _wordDocReader;
    private readonly FibReader _fib;
    private string _text = string.Empty;
    private List<Piece> _pieces = new();

    /// <param name="wordDocReader">Reader for the WordDocument stream</param>
    /// <param name="tableReader">Reader for the Table stream (0Table or 1Table)</param>
    /// <param name="fib">Parsed FIB</param>
    public TextReader(BinaryReader wordDocReader, BinaryReader tableReader, FibReader fib)
    {
        _wordDocReader = wordDocReader;
        TableReader = tableReader;
        _fib = fib;
    }

    /// <summary>
    /// Reader for the Table stream
    /// </summary>
    public BinaryReader TableReader { get; }

    /// <summary>
    /// Gets the complete reconstructed text of the main document body.
    /// </summary>
    public string Text => _text;

    /// <summary>
    /// Gets the piece table entries.
    /// </summary>
    public IReadOnlyList<Piece> Pieces => _pieces;

    /// <summary>
    /// Reads the complete text content from the main document body.
    /// </summary>
    public string ReadText()
    {
        if (_fib.CcpText <= 0) return string.Empty;

        if (_fib.FComplex)
        {
            ReadClxInternal();
            int totalCp = _fib.CcpText + _fib.CcpFtn + _fib.CcpHdd + _fib.CcpAtn + _fib.CcpEdn + _fib.CcpTxbx + _fib.CcpHdrTxbx;
            if (_pieces.Count > 0)
            {
                int maxCp = _pieces.Max(p => p.CpEnd);
                if (totalCp < maxCp)
                    totalCp = maxCp;
            }
            _text = ReconstructTextFromPieces(totalCp, null);
        }
        else
        {
            ReadClxInternal();
            if (_pieces.Count > 0)
            {
                int totalCp = _fib.CcpText + _fib.CcpFtn + _fib.CcpHdd + _fib.CcpAtn + _fib.CcpEdn + _fib.CcpTxbx + _fib.CcpHdrTxbx;
                int maxCp = _pieces.Max(p => p.CpEnd);
                if (totalCp < maxCp)
                    totalCp = maxCp;
                _text = ReconstructTextFromPieces(totalCp, null);
            }
            else
            {
                _text = ReadSimpleText();
            }
        }

        return _text;
    }

    /// <summary>
    /// Rebuilds document text from the piece table using optional per-CP CHP (for Lid-based encoding).
    /// Call ReadClx() and then FkpParser.ReadChpProperties() before this; pass the CHP map to use run-level Lid.
    /// </summary>
    public void SetTextFromPieces(int totalCpCount, IReadOnlyDictionary<int, ChpBase>? chpMap)
    {
        if (_pieces.Count == 0)
        {
            // use provided totalCpCount rather than default CcpText so that
            // simple documents with footnotes (where CcpFtn is zero) can still
            // include the extra characters that follow the main body.
            _text = ReadSimpleText(totalCpCount);
            return;
        }
        // the caller usually passes a totalCpCount calculated from FIB CP counts
        // (CcpText, CcpFtn, etc).  in a few buggy files – the sample that's been
        // giving us trouble for example – CcpFtn is zero even though the piece
        // table actually contains footnote text beyond the body.  when the
        // provided totalCpCount is too small we end up truncating the reconstructed
        // text and any later calls to GetText() return empty strings, which is why
        // our footnote reader produced empty notes.  to be robust we always make
        // sure the count is at least as large as the highest CP referenced in the
        // piece table; this allows us to ignore incorrect/zero CP limits.
        if (_pieces.Count > 0)
        {
            int maxCp = _pieces.Max(p => p.CpEnd);
            if (totalCpCount < maxCp)
                totalCpCount = maxCp;
        }
        _text = ReconstructTextFromPieces(totalCpCount, chpMap);
    }

    /// <summary>
    /// Gets text for a specific CP range.
    /// </summary>
    public string GetText(int startCp, int length)
    {
        if (string.IsNullOrEmpty(_text) || startCp >= _text.Length)
            return string.Empty;

        var end = Math.Min(startCp + length, _text.Length);
        return _text.Substring(startCp, end - startCp);
    }

    /// <summary>
    /// Decode a sequence of characters beginning at the specified file-character
    /// (FC) offset.  This is a thin wrapper around the piece-decoding helper
    /// so that callers such as <see cref="FootnoteReader"/> can bypass the
    /// global _text buffer when the CP→FC mapping is known.
    /// </summary>
    public string DecodeRangeFromFc(int fc, int charCount, ushort lid)
    {
        var fake = new Piece { FileOffset = (uint)fc, RawFcMasked = (uint)fc };
        return DecodeCompressedPieceWithLid(fake, charCount, lid);
    }

    /// <summary>
    /// Specialized helper for decoding a span of characters when the bytes
    /// are actually located in the Table stream rather than WordDocument.
    /// The FKP parser may return offsets that point at the 1Table/0Table
    /// stream (this happens when footnote text is stored there), so callers
    /// such as <see cref="FootnoteReader"/> can try this variant if the
    /// normal stream produce garbage.  The algorithm mirrors
    /// <see cref="DecodeCompressedPieceWithLid"/> but reads directly from
    /// <see cref="TableReader"/> and does not attempt the "raw vs fc"
    /// ambiguity since the caller already supplies the correct byte offset.
    /// </summary>
    public string DecodeRangeFromTableFc(int fc, int charCount, ushort lid)
    {
        if (charCount <= 0) return string.Empty;
        var stream = TableReader.BaseStream;
        int maxLen = (int)stream.Length;
        string bestStr = string.Empty;
        int bestScore = int.MinValue;

        // helper to score and update best
        void Consider(string s)
        {
            var q = DecodeQuality(s);
            if (q > bestScore) { bestScore = q; bestStr = s; }
        }

        // try Unicode interpretation (2 bytes per char)
        if (fc + charCount * 2 <= maxLen)
        {
            stream.Seek(fc, SeekOrigin.Begin);
            var buf = new byte[charCount * 2];
            stream.Read(buf, 0, buf.Length);
            Consider(Encoding.Unicode.GetString(buf));
        }

        // try ANSI interpretation using lid-dependent encoding
        if (fc + charCount <= maxLen)
        {
            stream.Seek(fc, SeekOrigin.Begin);
            var buf = new byte[charCount];
            stream.Read(buf, 0, buf.Length);
            var enc = GetEncodingForCompressedText(lid);
            try { Consider(enc.GetString(buf)); } catch { }

            // extra code pages as in DecodeCompressedPieceWithLid
            foreach (var extra in GetExtraEncodingsForCompressed(lid))
            {
                if (extra.CodePage == enc.CodePage) continue;
                try { Consider(extra.GetString(buf)); } catch { }
            }
        }

        return bestStr;
    }

    // ─── CLX Parsing ────────────────────────────────────────────────

    /// <summary>
    /// Reads the CLX structure from the Table stream (piece table).
    /// Call this before ReadChpProperties if using per-run Lid for decoding.
    /// </summary>
    public void ReadClx()
    {
        ReadClxInternal();
    }

    private void ReadClxInternal()
    {
        if (_fib.FcClx == 0 || _fib.LcbClx == 0)
            return;

        TableReader.BaseStream.Seek(_fib.FcClx, SeekOrigin.Begin);
        var endPosition = _fib.FcClx + _fib.LcbClx;

        // Skip any Prc entries (clxt = 0x01)
        while (TableReader.BaseStream.Position < endPosition)
        {
            var clxt = TableReader.ReadByte();

            if (clxt == 0x01)
            {
                // Prc — contains a GrpPrl
                var cbGrpprl = TableReader.ReadInt16();
                if (cbGrpprl > 0)
                    TableReader.BaseStream.Seek(cbGrpprl, SeekOrigin.Current);
            }
            else if (clxt == 0x02)
            {
                // Pcdt — the piece table
                var lcb = TableReader.ReadInt32(); // size of PlcPcd
                ReadPlcPcd(lcb);
                break;
            }
            else
            {
                // Unknown clxt — stop
                break;
            }
        }
    }

    /// <summary>
    /// Reads the PlcPcd (Piece Table) structure.
    /// 
    /// PlcPcd layout:
    ///   CP[0] CP[1] ... CP[n]       —  (n+1) × 4-byte CPs
    ///   PCD[0] PCD[1] ... PCD[n-1]  —  n × 8-byte Piece Descriptors
    ///
    /// where n = number of pieces.
    /// Total size = (n+1)*4 + n*8 = 4 + n*12
    /// So n = (lcb - 4) / 12
    /// </summary>
    private void ReadPlcPcd(int lcb)
    {
        if (lcb < 16) return; // minimum: 2 CPs + 1 PCD = 4+4+8 = 16

        var pieceCount = (lcb - 4) / 12;
        if (pieceCount <= 0) return;

        // Read CP array: (pieceCount + 1) entries
        var cps = new int[pieceCount + 1];
        for (int i = 0; i <= pieceCount; i++)
        {
            cps[i] = TableReader.ReadInt32();
        }

        // Read PCD array: pieceCount entries, each 8 bytes
        _pieces = new List<Piece>(pieceCount);
        for (int i = 0; i < pieceCount; i++)
        {
            var pcd = ReadPcd();
            var piece = new Piece
            {
                CpStart = cps[i],
                CpEnd = cps[i + 1],
                FileOffset = pcd.fc,
                RawFcMasked = pcd.rawFcMasked,
                IsUnicode = !pcd.fCompressed,
                Prm = pcd.prm
            };
            _pieces.Add(piece);
        }
    }

    /// <summary>
    /// Reads a single PCD (Piece Descriptor), 8 bytes:
    ///   ABCDxxxxh  (2 bytes) - first word, unused in practice
    ///   fc         (4 bytes) - file offset in WordDocument stream
    ///   prm        (2 bytes) - property modifier
    ///
    /// fc encoding (MS-DOC FcCompressed): low 30 bits = fc, bit 30 = fCompressed, bit 31 = reserved.
    ///   If fCompressed (bit 30) is set: text is ANSI, byte offset = fc/2.
    ///   If fCompressed is 0: text is UTF-16LE, byte offset = fc (30-bit value).
    /// </summary>
    private (uint fc, bool fCompressed, uint rawFcMasked, ushort prm) ReadPcd()
    {
        TableReader.ReadUInt16();
        var rawFc = TableReader.ReadUInt32();
        bool fCompressed = (rawFc & 0x40000000) != 0;
        var rawFcMasked = rawFc & 0x3FFFFFFFu; // 30-bit fc only
        uint fc = fCompressed ? rawFcMasked / 2 : rawFcMasked; // Unicode: byte offset = fc; Compressed: byte offset = fc/2
        var prm = TableReader.ReadUInt16();
        return (fc, fCompressed, rawFcMasked, prm);
    }

    // ─── Text Reconstruction ────────────────────────────────────────

    /// <summary>
    /// Scores how "valid" a decoded string looks. Penalizes replacement chars and mojibake-like
    /// symbols; rewards normal letters, digits, CJK. Used to pick the best decode for compressed pieces.
    /// </summary>
    private static int DecodeQuality(string s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        int score = 0;
        int cjkCount = 0;
        foreach (var c in s)
        {
            if (c == '\uFFFD') score -= 50;  // replacement char = strong mojibake signal
            else if (c >= '\u4E00' && c <= '\u9FFF') { score += 5; cjkCount++; }  // CJK unified
            else if (c >= '\u3400' && c <= '\u4DBF') { score += 5; cjkCount++; }  // CJK extension A
            else if (c >= '\u3000' && c <= '\u303F') score += 4;  // CJK symbols/punctuation
            else if (c >= '\u3040' && c <= '\u309F') { score += 4; cjkCount++; }  // Hiragana
            else if (c >= '\u30A0' && c <= '\u30FF') { score += 4; cjkCount++; }  // Katakana
            else if (c >= '\uAC00' && c <= '\uD7AF') { score += 4; cjkCount++; }  // Hangul syllables
            else if (char.IsLetterOrDigit(c) || c == ' ' || c == '\t' || c == '\r' || c == '\n') score += 2;
            else if (c >= 0x0B80 && c <= 0x0BFF) score -= 8;   // Malayalam (common in mojibake)
            else if (c >= 0x24B6 && c <= 0x24FF) score -= 5;   // enclosed alphanumerics
            else if (c >= 0x27F0 && c <= 0x27FF) score -= 5;   // supplementary arrows
            else if (c >= 0x00C0 && c <= 0x00FF && cjkCount > 0) score -= 3;  // Latin-1 accented when doc has CJK (likely wrong)
            else if (!char.IsControl(c)) score += 1;
        }
        // Bonus if string has CJK: prefer this decode when we're deciding between encodings
        if (cjkCount > 0) score += cjkCount;
        return score;
    }

    /// <summary>
    /// Gets the encoding for compressed (ANSI) text from the document's language ID (FIB lid).
    /// Note: Some Word versions store Lid=0x0409 (English) even for East Asian documents.
    /// </summary>
    private static Encoding GetEncodingForCompressedText(ushort lid)
    {
        try
        {
            if (lid == 0x0804 || lid == 0x0404) return Encoding.GetEncoding(936);   // Chinese Simplified/Traditional → GBK
            if (lid == 0x0411) return Encoding.GetEncoding(932);                    // Japanese → Shift-JIS
            if (lid == 0x0412) return Encoding.GetEncoding(949);                   // Korean → EUC-KR
        }
        catch (ArgumentException) { }
        return Encoding.GetEncoding(1252); // Western default
    }

    /// <summary>
    /// Tries to get additional encodings to try for compressed pieces (e.g. when Lid is wrong or mixed content).
    /// </summary>
    private static List<Encoding> GetExtraEncodingsForCompressed(ushort lid)
    {
        var list = new List<Encoding>();
        foreach (var cp in new[] { 936, 950, 54936, 65001, 932, 949 })  // GBK, Big5, GB18030, UTF-8, Shift-JIS, EUC-KR
        {
            try
            {
                var enc = Encoding.GetEncoding(cp);
                if (!list.Any(e => e.CodePage == enc.CodePage))
                    list.Add(enc);
            }
            catch (ArgumentException) { }
        }
        if (lid != 0x0804 && lid != 0x0404 && lid != 0x0411 && lid != 0x0412)
        {
            try { list.Add(Encoding.GetEncoding(1252)); } catch { }
        }
        return list;
    }

    /// <summary>
    /// Reconstructs the main document text from piece table entries.
    /// When chpMap is provided, uses per-run Lid for compressed pieces (fixes mixed-language partial mojibake).
    /// </summary>
    private string ReconstructTextFromPieces(int totalCpCount, IReadOnlyDictionary<int, ChpBase>? chpMap)
    {
        if (_pieces.Count == 0) return string.Empty;

        var sb = new StringBuilder(Math.Max(0, totalCpCount));

        foreach (var piece in _pieces)
        {
            var cpStart = piece.CpStart;
            var cpEnd = Math.Min(piece.CpEnd, totalCpCount);
            if (cpStart >= cpEnd) continue;
            if (cpStart >= totalCpCount) break;

            var charCount = cpEnd - cpStart;

            if (piece.IsUnicode)
            {
                var byteCount = charCount * 2;
                if (piece.FileOffset + byteCount > _wordDocReader.BaseStream.Length)
                    byteCount = (int)Math.Max(0, _wordDocReader.BaseStream.Length - piece.FileOffset);
                if (byteCount <= 0) continue;
                byteCount &= ~1;
                _wordDocReader.BaseStream.Seek(piece.FileOffset, SeekOrigin.Begin);
                var bytes = _wordDocReader.ReadBytes(byteCount);
                if (bytes.Length >= 2)
                    sb.Append(Encoding.Unicode.GetString(bytes, 0, bytes.Length - (bytes.Length % 2)));
            }
            else
            {
                string pieceText = DecodeCompressedPiece(piece, cpStart, cpEnd, charCount, chpMap);
                sb.Append(pieceText);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Decodes a single compressed piece, optionally per-run by chpMap Lid to fix partial mojibake.
    /// </summary>
    private string DecodeCompressedPiece(Piece piece, int cpStart, int cpEnd, int charCount, IReadOnlyDictionary<int, ChpBase>? chpMap)
    {
        // Run-level decoding: when chpMap has boundaries inside this piece, decode each run with its own Lid
        if (chpMap != null && chpMap.Count > 0)
        {
            var boundaries = chpMap.Keys.Where(cp => cp > cpStart && cp < cpEnd).OrderBy(cp => cp).ToList();
            if (boundaries.Count > 0)
            {
                var sb = new StringBuilder(charCount);
                int runStart = cpStart;
                foreach (var runEndCp in boundaries)
                {
                    var runLen = runEndCp - runStart;
                    if (runLen <= 0) continue;
                    ushort lid = _fib.Lid;
                    if (chpMap.TryGetValue(runStart, out var chp) && chp.Language != 0)
                        lid = (ushort)(chp.Language & 0xFFFF);
                    var runBytes = ReadCompressedPieceSlice(piece, runStart - cpStart, runLen);
                    sb.Append(DecodeCompressedBytes(runBytes, lid));
                    runStart = runEndCp;
                }
                if (runStart < cpEnd)
                {
                    ushort lid = _fib.Lid;
                    if (chpMap.TryGetValue(runStart, out var chp) && chp.Language != 0)
                        lid = (ushort)(chp.Language & 0xFFFF);
                    var runBytes = ReadCompressedPieceSlice(piece, runStart - cpStart, cpEnd - runStart);
                    sb.Append(DecodeCompressedBytes(runBytes, lid));
                }
                return sb.ToString();
            }
        }

        // Piece-level: single Lid at piece start
        ushort pieceLid = _fib.Lid;
        if (chpMap != null && chpMap.TryGetValue(cpStart, out var pieceChp) && pieceChp.Language != 0)
            pieceLid = (ushort)(pieceChp.Language & 0xFFFF);
        return DecodeCompressedPieceWithLid(piece, charCount, pieceLid);
    }

    private byte[] ReadCompressedPieceSlice(Piece piece, int byteOffset, int byteCount)
    {
        if (byteCount <= 0) return Array.Empty<byte>();
        var streamOffset = piece.FileOffset + (uint)byteOffset;
        if (streamOffset + byteCount > _wordDocReader.BaseStream.Length)
            byteCount = (int)Math.Max(0, _wordDocReader.BaseStream.Length - streamOffset);
        if (byteCount <= 0) return Array.Empty<byte>();
        _wordDocReader.BaseStream.Seek(streamOffset, SeekOrigin.Begin);
        return _wordDocReader.ReadBytes(byteCount);
    }

    /// <summary>
    /// Tries multiple encodings and offsets for a compressed piece; prefers Unicode when ANSI decode is poor.
    /// </summary>
    private string DecodeCompressedPieceWithLid(Piece piece, int charCount, ushort lid)
    {
        var bestStr = "";
        var bestScore = int.MinValue;
        string? unicodeStrAtFc = null;
        int unicodeScoreAtFc = int.MinValue;
        string? unicodeStrAtRaw = null;
        int unicodeScoreAtRaw = int.MinValue;
        byte[] ansiBytes = Array.Empty<byte>();

        // 1. Unicode at RawFcMasked
        _wordDocReader.BaseStream.Seek(piece.RawFcMasked, SeekOrigin.Begin);
        if (piece.RawFcMasked + (uint)(charCount * 2) <= _wordDocReader.BaseStream.Length)
        {
            var s1 = Encoding.Unicode.GetString(_wordDocReader.ReadBytes(charCount * 2), 0, charCount * 2);
            var q1 = DecodeQuality(s1);
            unicodeStrAtRaw = s1;
            unicodeScoreAtRaw = q1;
            if (q1 > bestScore) { bestScore = q1; bestStr = s1; }
        }

        // 2. Unicode at FileOffset (standard for "compressed" flag with UTF-16 data)
        _wordDocReader.BaseStream.Seek(piece.FileOffset, SeekOrigin.Begin);
        if (piece.FileOffset + (uint)(charCount * 2) <= _wordDocReader.BaseStream.Length)
        {
            var s2 = Encoding.Unicode.GetString(_wordDocReader.ReadBytes(charCount * 2), 0, charCount * 2);
            var q2 = DecodeQuality(s2);
            unicodeStrAtFc = s2;
            unicodeScoreAtFc = q2;
            if (q2 > bestScore) { bestScore = q2; bestStr = s2; }
        }

        // 3. ANSI at FileOffset
        _wordDocReader.BaseStream.Seek(piece.FileOffset, SeekOrigin.Begin);
        ansiBytes = _wordDocReader.ReadBytes(charCount);
        var encLid = GetEncodingForCompressedText(lid);
        var s3 = encLid.GetString(ansiBytes);
        var q3 = DecodeQuality(s3);
        if (q3 > bestScore) { bestScore = q3; bestStr = s3; }

        // 4. ANSI at RawFcMasked
        if (piece.RawFcMasked + (uint)charCount <= _wordDocReader.BaseStream.Length)
        {
            _wordDocReader.BaseStream.Seek(piece.RawFcMasked, SeekOrigin.Begin);
            var ansiBytesAlt = _wordDocReader.ReadBytes(charCount);
            var s4 = encLid.GetString(ansiBytesAlt);
            var q4 = DecodeQuality(s4);
            if (q4 > bestScore) { bestScore = q4; bestStr = s4; ansiBytes = ansiBytesAlt; }
        }

        // 5. Extra code pages on the best ansi bytes we have
        foreach (var enc in GetExtraEncodingsForCompressed(lid))
        {
            if (enc.CodePage == encLid.CodePage) continue;
            try
            {
                var s = enc.GetString(ansiBytes);
                var q = DecodeQuality(s);
                if (q > bestScore) { bestScore = q; bestStr = s; }
            }
            catch { /* ignore */ }
        }

        // Prefer Unicode when ANSI result is clearly bad (replacement chars or very low score)
        bool ansiHasReplacement = bestStr.Contains('\uFFFD');
        bool ansiLowQuality = bestScore < 0 || (bestStr.Length > 0 && bestScore < bestStr.Length);
        var bestUnicodeScore = Math.Max(unicodeScoreAtFc, unicodeScoreAtRaw);
        var bestUnicodeStr = unicodeScoreAtFc >= unicodeScoreAtRaw ? unicodeStrAtFc : unicodeStrAtRaw;
        if (bestUnicodeStr != null && !bestUnicodeStr.Contains('\uFFFD') && bestUnicodeScore > 0 &&
            (ansiHasReplacement || (ansiLowQuality && bestUnicodeScore >= bestScore)))
            bestStr = bestUnicodeStr;

        if (string.IsNullOrEmpty(bestStr) && ansiBytes.Length > 0)
            bestStr = encLid.GetString(ansiBytes);
        return bestStr;
    }

    /// <summary>
    /// Decodes a byte slice as compressed (ANSI) text with given Lid; tries multiple encodings.
    /// </summary>
    private string DecodeCompressedBytes(byte[] bytes, ushort lid)
    {
        if (bytes.Length == 0) return string.Empty;
        var encLid = GetEncodingForCompressedText(lid);
        var bestStr = encLid.GetString(bytes);
        var bestScore = DecodeQuality(bestStr);
        foreach (var enc in GetExtraEncodingsForCompressed(lid))
        {
            if (enc.CodePage == encLid.CodePage) continue;
            try
            {
                var s = enc.GetString(bytes);
                var q = DecodeQuality(s);
                if (q > bestScore) { bestScore = q; bestStr = s; }
            }
            catch { /* ignore */ }
        }
        return bestStr;
    }

    /// <summary>
    /// Fallback: reads text directly from WordDocument stream for non-complex documents
    /// that have no CLX. This is rare for Word 97+ but handled for safety.
    /// </summary>
    private string ReadSimpleText()
    {
        return ReadSimpleText(_fib.CcpText);
    }

    /// <summary>
    /// Read simple text from the WordDocument stream using a specific CP count.
    /// Used when the caller knows the story length may exceed the FIB's CcpText
    /// (e.g. footnotes stored after the main body when CcpFtn is zero).
    /// </summary>
    private string ReadSimpleText(int cpCount)
    {
        // Improved simple-text reader that tries multiple encodings just like
        // DecodeCompressedPieceWithLid does for pieces.  The previous version
        // assumed Unicode (2 bytes/char) and only fell back to ANSI when the
        // decoded string was half nulls; this frequently mis‑decoded footnotes
        // and other stories in "simple" documents, turning valid bytes into
        // unpaired surrogates which our sanitizer then discarded.  By reusing
        // the same heuristics as the piece decoder we can correctly handle
        // ANSI/GBK, Unicode, broken FIB counts, etc.

        if (cpCount <= 0) return string.Empty;
        var textOffset = _fib.FcMin > 0 ? (int)_fib.FcMin : 0x200;

        // Build a temporary Piece object so we can reuse the decoding helper.
        var fake = new Piece
        {
            FileOffset = (uint)textOffset,
            RawFcMasked = (uint)textOffset
        };
        return DecodeCompressedPieceWithLid(fake, cpCount, _fib.Lid);
    }
}

/// <summary>
/// Represents a piece in the Piece Table.
/// Each piece maps a range of Character Positions (CPs) to a byte offset
/// in the WordDocument stream.
/// </summary>
public class Piece
{
    /// <summary>Starting CP (inclusive)</summary>
    public int CpStart { get; set; }

    /// <summary>Ending CP (exclusive)</summary>
    public int CpEnd { get; set; }

    /// <summary>Byte offset in the WordDocument stream</summary>
    public uint FileOffset { get; set; }

    /// <summary>Raw PCD fc with bit 30 masked (for alternate offset when re-decoding compressed as Unicode)</summary>
    public uint RawFcMasked { get; set; }

    /// <summary>True if text at this offset is Unicode (UTF-16LE), false if ANSI</summary>
    public bool IsUnicode { get; set; }

    /// <summary>Property modifier (Prm) from the PCD</summary>
    public ushort Prm { get; set; }

    /// <summary>Number of characters in this piece</summary>
    public int CharCount => CpEnd - CpStart;

    // Legacy compatibility properties
    public int Start { get => CpStart; set => CpStart = value; }
    public int End { get => CpEnd; set => CpEnd = value; }
    public int Length => CharCount;
}

/// <summary>
/// Document Properties (DOP) Reader.
/// Reads from the Table stream at fcDop offset.
/// </summary>
public class DocumentPropertiesReader
{
    private readonly BinaryReader TableReader;
    private readonly FibReader _fib;

    public DocumentPropertiesReader(BinaryReader tableReader, FibReader fib)
    {
        TableReader = tableReader;
        _fib = fib;
    }

    /// <summary>
    /// Reads document properties from the Table stream.
    /// </summary>
    public DocumentProperties Read()
    {
        var props = new DocumentProperties();

        if (_fib.FcDop == 0 || _fib.LcbDop == 0) return props;

        TableReader.BaseStream.Seek(_fib.FcDop, SeekOrigin.Begin);

        // DOP is a variable-length structure. We read known fields.
        // Per MS-DOC §2.7.4 (Dop97):
        // The DOP structure has evolved across Word versions.
        // We read the common fields carefully.

        var dopBytes = TableReader.ReadBytes((int)Math.Min(_fib.LcbDop, 500));
        if (dopBytes.Length < 20) return props;

        // Offsets within DOP (Dop97 structure):
        // 0-1: bit flags (fWidowControl, fPaginated, fFacingPages, etc.)
        // 2-3: bit flags continued
        // 4-5: bit flags continued
        // 6-7: bit flags continued
        // 14-15: dxaTab (default tab width)
        // 16-17: dxaColumns (column width)
        // 28-29: itxtWrap (text wrapping)
        // For page setup, we rely on section properties (SEP) instead of DOP.
        // DOP primarily stores document-level flags.

        // Extract bit flags from Dop97 structure
        if (dopBytes.Length >= 4)
        {
            var flags0 = BitConverter.ToUInt16(dopBytes, 0);
            var flags1 = BitConverter.ToUInt16(dopBytes, 2);

            // Group 0 flags (from MS-DOC §2.7.4)
            props.FWidowControl = (flags0 & 0x0002) != 0;      // fWidowControl
            props.FPaginated = (flags0 & 0x0004) != 0;         // fPaginated
            props.FFacingPages = (flags0 & 0x0008) != 0;       // fFacingPages
            props.FBreaks = (flags0 & 0x0010) != 0;            // fBreaks
            props.FAutoHyphenate = (flags0 & 0x0020) != 0;     // fAutoHyphenate
            props.FDoHyphenation = (flags0 & 0x0040) != 0;     // fDoHyphenation
            props.FFELayout = (flags0 & 0x0080) != 0;          // fFELayout
            props.FLayoutSameAsWin95 = (flags0 & 0x0100) != 0; // fLayoutSameAsWin95
            props.FPrintBodyBeforeHeaders = (flags0 & 0x0200) != 0; // fPrintBodyBeforeHeaders
            props.FSuppressBottomSpacing = (flags0 & 0x0400) != 0; // fSuppressBottomSpacing
            props.FWrapAuto = (flags0 & 0x0800) != 0;          // fWrapAuto
            props.FPrintPaperBefore = (flags0 & 0x1000) != 0;  // fPrintPaperBefore
            props.FSuppressSpacings = (flags0 & 0x2000) != 0;  // fSuppressSpacings
            props.FMirrorMargins = (flags0 & 0x4000) != 0;     // fMirrorMargins
            // bits 14-15: fRuler, fNoTabForInd

            // Group 1 flags
            props.FUsePrinterMetrics = (flags1 & 0x0001) != 0; // fUsePrinterMetrics
            props.FNoPgp = (flags1 & 0x0002) != 0;             // fNoPgp
            props.FShrinkToFit = (flags1 & 0x0004) != 0;       // fShrinkToFit
            props.FPrintFormsData = (flags1 & 0x0008) != 0;    // fPrintFormsData
            props.FAllowPositionOnOnly = (flags1 & 0x0010) != 0; // fAllowPositionOnOnly
            props.FDisplayBackground = (flags1 & 0x0020) != 0; // fDisplayBackground
            props.FDisplayLineNumbers = (flags1 & 0x0040) != 0; // fDisplayLineNumbers
            props.FPrintMicros = (flags1 & 0x0080) != 0;       // fPrintMicros
            props.FSaveFormsData = (flags1 & 0x0100) != 0;     // fSaveFormsData
            props.FDisplayColBreak = (flags1 & 0x0200) != 0;   // fDisplayColBreak
            props.FDisplayPageEnd = (flags1 & 0x0400) != 0;    // fDisplayPageEnd
            props.FDisplayUnits = (flags1 & 0x0800) != 0;      // fDisplayUnits
            props.FProtectForms = (flags1 & 0x1000) != 0;      // fProtectForms
            props.FProtectSparce = (flags1 & 0x2000) != 0;     // fProtectSparce
            props.FConsecutiveHyphen = (flags1 & 0x4000) != 0; // fConsecutiveHyphen
            props.FLetterFinal = (flags1 & 0x8000) != 0;       // fLetterFinal
        }

        if (dopBytes.Length >= 16)
        {
            // dxaTab (default tab width) at offset 14-15
            props.DxaTab = BitConverter.ToInt16(dopBytes, 14);
            // dxaColumns (column width) at offset 16-17
            props.DxaColumns = BitConverter.ToInt16(dopBytes, 16);
        }

        if (dopBytes.Length >= 30)
        {
            // itxtWrap at offset 28-29
            props.ITxtWrap = BitConverter.ToInt16(dopBytes, 28);
        }

        // Default margins and page size will come from section properties
        // For now, return defaults that are reasonable
        props.PageWidth = 12240;   // 8.5" in twips
        props.PageHeight = 15840;  // 11" in twips
        props.MarginTop = 1440;    // 1" in twips
        props.MarginBottom = 1440;
        props.MarginLeft = 1800;   // 1.25" in twips
        props.MarginRight = 1800;

        return props;
    }
}
