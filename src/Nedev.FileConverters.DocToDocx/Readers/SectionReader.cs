using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Reads Section Properties (SEPX) from the DOC binary format.
/// MS-DOC §2.6.2 Section Properties.
/// </summary>
public class SectionReader
{
    private readonly BinaryReader _tableReader;
    private readonly BinaryReader _wordDocReader;
    private readonly FibReader _fib;

    public SectionReader(BinaryReader tableReader, BinaryReader wordDocReader, FibReader fib)
    {
        _tableReader = tableReader;
        _wordDocReader = wordDocReader;
        _fib = fib;
    }

    /// <summary>
    /// Reads all sections defined in the PlcfSed.
    /// </summary>
    public List<SectionInfo> ReadSections()
    {
        var sections = new List<SectionInfo>();
        if (_fib.FcPlcfSed == 0 || _fib.LcbPlcfSed == 0) return sections;

        if (!_tableReader.CanReadRange(_fib.FcPlcfSed, _fib.LcbPlcfSed) || _fib.LcbPlcfSed < 20)
        {
            Logger.Warning($"Skipped section PLC parsing because PlcfSed range 0x{_fib.FcPlcfSed:X}/0x{_fib.LcbPlcfSed:X} is invalid.");
            return sections;
        }

        _tableReader.BaseStream.Seek(_fib.FcPlcfSed, SeekOrigin.Begin);
        
        // n = number of SED structures
        // lcbPlcfSed = (n+1)*4 + n*cbSed
        // cbSed = 12 bytes
        int n = (int)((_fib.LcbPlcfSed - 4) / 16);
        if (n <= 0) return sections;

        int[] cps = new int[n + 1];
        for (int i = 0; i <= n; i++)
        {
            cps[i] = _tableReader.ReadInt32();
        }

        for (int i = 0; i < n; i++)
        {
            // Read SED (12 bytes)
            // SED layout is 12 bytes. The SEPX offset is not stored at byte 0;
            // there is a 2-byte field ahead of it, so reading the first four
            // bytes as fcSepx misaligns the value and produces bogus high-bit
            // offsets like 0x14000036.
            _tableReader.ReadUInt16();
            uint fcSepx = _tableReader.ReadUInt32();
            _tableReader.BaseStream.Seek(6, SeekOrigin.Current);

            var section = new SectionInfo
            {
                StartCp = cps[i],
                EndCp = cps[i + 1]
            };

            // Seed each section with the SPRM parser's default SEP values so
            // invalid or missing SEPX records still fall back to Word-like page
            // geometry instead of the broader DocumentModel defaults.
            MapSepToSectionInfo(new SepBase(), section);

            if (fcSepx != 0xFFFFFFFF)
            {
                // Read SEPX from WordDocument stream, but guard against bogus offsets
                long originalWordPos = _wordDocReader.BaseStream.Position;
                try
                {
                    // make sure offset is within the stream
                    if (fcSepx < 0 || fcSepx + 2 > _wordDocReader.BaseStream.Length)
                        throw new IOException("SEPX offset outside WordDocument stream");

                    _wordDocReader.BaseStream.Seek(fcSepx, SeekOrigin.Begin);
                    ushort cb = _wordDocReader.ReadUInt16();

                    // sanity check the claimed length before reading bytes
                    if (cb > 0)
                    {
                        if (fcSepx + 2 + cb > _wordDocReader.BaseStream.Length)
                            throw new IOException($"SEPX size {cb} extends past end of stream");

                        byte[] grpprl = _wordDocReader.ReadBytes(cb);
                        var sepBase = new SepBase();
                        var parser = new SprmParser(_wordDocReader, 0);
                        parser.ApplyToSep(grpprl, sepBase);
                        
                        MapSepToSectionInfo(sepBase, section);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to read SEPX at 0x{fcSepx:X}; continuing with default section properties.", ex);
                }
                finally
                {
                    _wordDocReader.BaseStream.Seek(originalWordPos, SeekOrigin.Begin);
                }
            }

            sections.Add(section);
        }

        return sections;
    }

    private void MapSepToSectionInfo(SepBase sep, SectionInfo info)
    {
        info.PageWidth = sep.PageWidth;
        info.PageHeight = sep.PageHeight;
        info.MarginLeft = sep.MarginLeft;
        info.MarginRight = sep.MarginRight;
        info.MarginTop = sep.MarginTop;
        info.MarginBottom = sep.MarginBottom;
        info.HeaderMargin = sep.MarginHeader;
        info.FooterMargin = sep.MarginFooter;
        info.Gutter = sep.Gutter;
        info.ColumnCount = sep.ColumnCount;
        info.ColumnSpacing = sep.ColumnSpacing;
        info.DocGridLinePitch = sep.DocGridLinePitch;
        info.BreakCode = sep.BreakCode;
        info.VerticalAlignment = sep.VerticalAlignment;
        
        // Section setup often defaults to Portrait unless orientation is flipped
        info.IsLandscape = (sep.PageWidth > sep.PageHeight);
    }
}
