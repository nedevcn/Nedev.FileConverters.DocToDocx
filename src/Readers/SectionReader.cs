using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Readers;

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
            // fcSepx (4 bytes): offset in WordDocument stream to the SEPX
            uint fcSepx = _tableReader.ReadUInt32();
            _tableReader.BaseStream.Seek(8, SeekOrigin.Current); // Reserved

            var section = new SectionInfo
            {
                StartCp = cps[i],
                EndCp = cps[i + 1]
            };

            if (fcSepx != 0xFFFFFFFF)
            {
                // Read SEPX from WordDocument stream
                long originalWordPos = _wordDocReader.BaseStream.Position;
                try
                {
                    _wordDocReader.BaseStream.Seek(fcSepx, SeekOrigin.Begin);
                    ushort cb = _wordDocReader.ReadUInt16();
                    if (cb > 0)
                    {
                        byte[] grpprl = _wordDocReader.ReadBytes(cb);
                        var sepBase = new SepBase();
                        var parser = new SprmParser(_wordDocReader, 0);
                        parser.ApplyToSep(grpprl, sepBase);
                        
                        MapSepToSectionInfo(sepBase, section);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to read SEPX at 0x{fcSepx:X}: {ex.Message}");
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
        info.BreakCode = sep.BreakCode;
        info.VerticalAlignment = sep.VerticalAlignment;
        
        // Section setup often defaults to Portrait unless orientation is flipped
        info.IsLandscape = (sep.PageWidth > sep.PageHeight);
    }
}
