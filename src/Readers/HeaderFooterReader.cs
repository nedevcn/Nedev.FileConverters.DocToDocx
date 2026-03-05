using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Header/Footer reader - parses PlcfHdd structure from the Table stream.
///
/// Based on MS-DOC specification §2.7.
///
/// PlcfHdd structure contains:
///   - Array of CP (character position) boundaries
///   - Array of Hdd (header/footer descriptor) entries
///
/// Each section can have up to 6 header/footer types:
///   - First page header/footer
///   - Odd page header/footer (default)
///   - Even page header/footer
/// </summary>
public class HeaderFooterReader
{
    private readonly BinaryReader _tableReader;
    private readonly BinaryReader _wordDocReader;
    private readonly FibReader _fib;
    private readonly TextReader _textReader;

    public List<HeaderFooterModel> Headers { get; private set; } = new();
    public List<HeaderFooterModel> Footers { get; private set; } = new();

    public HeaderFooterReader(
        BinaryReader tableReader,
        BinaryReader wordDocReader,
        FibReader fib,
        TextReader textReader)
    {
        _tableReader = tableReader;
        _wordDocReader = wordDocReader;
        _fib = fib;
        _textReader = textReader;
    }

    /// <summary>
    /// Reads header/footer information from the document.
    /// </summary>
    public void Read(DocumentModel document)
    {
        if (_fib.FcPlcfHdd == 0 || _fib.LcbPlcfHdd == 0)
        {
            // No header/footer data
            return;
        }

        try
        {
            ReadPlcfHdd(document);
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read headers/footers", ex);
        }
    }

    private void ReadPlcfHdd(DocumentModel document)
    {
        _tableReader.BaseStream.Seek(_fib.FcPlcfHdd, SeekOrigin.Begin);

        var dataSize = (int)_fib.LcbPlcfHdd;
        if (dataSize < 4) return;

        // PlcfHdd structure consists solely of an array of CPs.
        // n = number of header/footer boundaries. Each CP is 4 bytes.
        // Total dataSize = (n + 1) * 4
        var entryCount = (dataSize - 4) / 4;
        if (entryCount <= 0) return;

        var cpArray = new int[entryCount + 1];
        for (int i = 0; i <= entryCount; i++)
        {
            cpArray[i] = _tableReader.ReadInt32();
        }

        // The entries are grouped into blocks of 6 for each section:
        // 0: Header Even
        // 1: Header Odd (also used as default if no even/first exists)
        // 2: Footer Even
        // 3: Footer Odd (also used as default if no even/first exists)
        // 4: Header First
        // 5: Footer First
        int sectionCount = entryCount / 6;

        for (int sectionIndex = 0; sectionIndex < sectionCount; sectionIndex++)
        {
            int baseIdx = sectionIndex * 6;
            
            var types = new[]
            {
                HeaderFooterType.HeaderEven,
                HeaderFooterType.HeaderOdd,
                HeaderFooterType.FooterEven,
                HeaderFooterType.FooterOdd,
                HeaderFooterType.HeaderFirst,
                HeaderFooterType.FooterFirst
            };

            for (int t = 0; t < 6; t++)
            {
                int currentCp = cpArray[baseIdx + t];
                int nextCp = cpArray[baseIdx + t + 1];
                int length = nextCp - currentCp;

                if (length <= 0) continue;

                var type = types[t];

                try
                {
                    // Extract header/footer text
                    var text = ExtractHeaderFooterText(currentCp, length);

                    var model = new HeaderFooterModel
                    {
                        Type = type,
                        SectionIndex = sectionIndex,
                        Text = text,
                        CharacterPosition = currentCp,
                        CharacterLength = length
                    };

                    if (type == HeaderFooterType.HeaderFirst ||
                        type == HeaderFooterType.HeaderOdd ||
                        type == HeaderFooterType.HeaderEven)
                    {
                        Headers.Add(model);
                    }
                    else
                    {
                        Footers.Add(model);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to extract header/footer type {type} for section {sectionIndex}", ex);
                }
            }
        }
    }

    /// <summary>
    /// Extracts header/footer text from the global text stream.
    /// Header/footer text is stored after the main document and footnote text.
    /// </summary>
    private string ExtractHeaderFooterText(int cp, int length)
    {
        if (length <= 0)
            return string.Empty;

        // The header/footer text starts at CP position (CcpText + CcpFtn) in the global stream
        int headerStoryStartCp = _fib.CcpText + _fib.CcpFtn;
        int absoluteCp = headerStoryStartCp + cp;

        string rawText = _textReader.GetText(absoluteCp, length);
        return CleanHeaderFooterText(rawText);
    }

    /// <summary>
    /// Cleans header/footer text by removing control characters.
    /// </summary>
    private string CleanHeaderFooterText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            // Skip invalid XML characters (0x00-0x1F except tab, newline, carriage return)
            if (ch < 0x09 || (ch > 0x0D && ch < 0x20))
            {
                continue;
            }
            // Skip special Word characters
            switch (ch)
            {
                case '\x01':  // Field begin mark
                case '\x13': // Field separator
                case '\x14': // Field end
                case '\x15': // Object anchor
                    continue;
                case '\x0B':
                    sb.Append('\n');
                    break;
                case '\x07':
                    sb.Append('\t');
                    break;
                case '\x1E':
                    sb.Append('-');
                    break;
                case '\x1F':
                    break;
                default:
                    sb.Append(ch);
                    break;
            }
        }

        return sb.ToString().Trim();
    }

    /// <summary>
    /// Gets headers for a specific section.
    /// </summary>
    public List<HeaderFooterModel> GetHeadersForSection(int sectionIndex)
    {
        return Headers.Where(h => h.SectionIndex == sectionIndex).ToList();
    }

    /// <summary>
    /// Gets footers for a specific section.
    /// </summary>
    public List<HeaderFooterModel> GetFootersForSection(int sectionIndex)
    {
        return Footers.Where(f => f.SectionIndex == sectionIndex).ToList();
    }
}


