using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class ListReader
{
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;

    public List<NumberingDefinition> NumberingDefinitions { get; private set; } = new();
    public List<ListFormat> ListFormats { get; private set; } = new();

    public ListReader(BinaryReader tableReader, FibReader fib)
    {
        _tableReader = tableReader;
        _fib = fib;
    }

    public void Read()
    {
        if (_fib.FcPlcfLst == 0 || _fib.LcbPlcfLst == 0)
        {
            return;
        }

        try
        {
            ReadPlcfLst();
            ReadPlfLfo();
            BuildNumberingDefinitions();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Failed to read list formats: {ex.Message}");
        }
    }

    private void BuildNumberingDefinitions()
    {
        NumberingDefinitions = new List<NumberingDefinition>();

        foreach (var listFormat in ListFormats)
        {
            var numDef = new NumberingDefinition
            {
                Id = listFormat.ListId
            };

            for (int i = 0; i < Math.Min(listFormat.Levels.Count, 9); i++)
            {
                var lvl = listFormat.Levels[i];
                var numLevel = new NumberingLevel
                {
                    Level = lvl.Level,
                    NumberFormat = lvl.NumberFormat,
                    Text = lvl.NumberText ?? "%" + (i + 1),
                    Start = lvl.StartAt,
                    ParagraphProperties = lvl.ParagraphProperties,
                    RunProperties = lvl.RunProperties
                };
                numDef.Levels.Add(numLevel);
            }

            if (numDef.Levels.Count == 0)
            {
                for (int i = 0; i < 9; i++)
                {
                    numDef.Levels.Add(new NumberingLevel
                    {
                        Level = i,
                        NumberFormat = NumberFormat.Decimal,
                        Text = "%" + (i + 1),
                        Start = 1
                    });
                }
            }

            NumberingDefinitions.Add(numDef);
        }
    }

    private void ReadPlcfLst()
    {
        _tableReader.BaseStream.Seek(_fib.FcPlcfLst, SeekOrigin.Begin);
        var endPos = _fib.FcPlcfLst + _fib.LcbPlcfLst;

        if (endPos - _tableReader.BaseStream.Position < 4)
            return;

        var lstfCount = _tableReader.ReadInt32();
        if (lstfCount <= 0 || lstfCount > 1000)
            lstfCount = 64;

        // Phase 1: Read all LSTF headers (28 bytes each)
        var lstfHeaders = new List<(int lsid, ListType type)>();
        for (int i = 0; i < lstfCount && _tableReader.BaseStream.Position + 28 <= endPos; i++)
        {
            try
            {
                var lsid = _tableReader.ReadInt32(); // lsid
                _tableReader.ReadUInt32(); // tplc
                // rgistd[9] = 18 bytes
                for (int j = 0; j < 9; j++)
                    _tableReader.ReadUInt16();
                var flags = _tableReader.ReadUInt16();
                var listType = (ListType)(flags & 0x03);
                lstfHeaders.Add((lsid, listType));
            }
            catch
            {
                break;
            }
        }

        // Phase 2: Read LVLF structures that follow the LSTF array
        var lists = new List<ListFormat>();
        foreach (var (lsid, listType) in lstfHeaders)
        {
            if (lsid == 0) continue;
            var listFormat = new ListFormat
            {
                ListId = lsid,
                Type = listType
            };

            // Each LSTF has 9 levels (or 1 for simple lists, but we read 9 anyway)
            int levelCount = listType == ListType.Simple ? 1 : 9;
            for (int lvl = 0; lvl < levelCount; lvl++)
            {
                if (_tableReader.BaseStream.Position + 28 > endPos)
                {
                    // Not enough data for LVLF, create default level
                    listFormat.Levels.Add(CreateDefaultLevel(lvl, listType));
                    continue;
                }

                try
                {
                    var level = ReadLvlf(lvl, listType, endPos);
                    listFormat.Levels.Add(level);
                }
                catch
                {
                    listFormat.Levels.Add(CreateDefaultLevel(lvl, listType));
                }
            }

            // Fill remaining levels for simple lists
            for (int lvl = levelCount; lvl < 9; lvl++)
            {
                listFormat.Levels.Add(CreateDefaultLevel(lvl, listType));
            }

            lists.Add(listFormat);
        }

        ListFormats = lists;
    }

    /// <summary>
    /// Reads a single LVLF (List Level Format) structure.
    /// LVLF layout (28 bytes minimum):
    ///   iStartAt (4 bytes) - Starting number
    ///   nfc      (1 byte)  - Number format code  
    ///   jc       (1 byte)  - Alignment (bits 0-1)
    ///   flags    (1 byte)  - Various flags
    ///   flags2   (1 byte)  - More flags
    ///   rgbxchNums[9] (9 bytes) - Placeholder positions in xst
    ///   ixchFollow (1 byte) - What follows the number (tab, space, nothing)
    ///   dxaIndentSav (4 bytes) - Saved indent
    ///   reserved (4 bytes)
    ///   cbGrpprlChpx (1 byte)  - Size of CHP SPRM data
    ///   cbGrpprlPapx (1 byte)  - Size of PAP SPRM data
    ///   ilvlRestartLim (1 byte)
    ///   grfhic (1 byte)
    /// Followed by:
    ///   grpprlPapx (cbGrpprlPapx bytes) - PAP SPRMs
    ///   grpprlChpx (cbGrpprlChpx bytes) - CHP SPRMs
    ///   xst string (2 + len*2 bytes)    - Number text
    /// </summary>
    private ListLevel ReadLvlf(int levelIndex, ListType listType, long endPos)
    {
        var level = new ListLevel { Level = levelIndex };

        // Read LVLF fixed portion (28 bytes)
        level.StartAt = _tableReader.ReadInt32();        // iStartAt
        byte nfc = _tableReader.ReadByte();              // nfc
        byte jcByte = _tableReader.ReadByte();           // jc (bits 0-1)
        _tableReader.ReadByte();                         // flags
        _tableReader.ReadByte();                         // flags2
        var rgbxchNums = _tableReader.ReadBytes(9);      // rgbxchNums
        _tableReader.ReadByte();                         // ixchFollow
        _tableReader.ReadInt32();                        // dxaIndentSav
        _tableReader.ReadInt32();                        // reserved
        byte cbGrpprlChpx = _tableReader.ReadByte();
        byte cbGrpprlPapx = _tableReader.ReadByte();
        _tableReader.ReadByte();                         // ilvlRestartLim
        _tableReader.ReadByte();                         // grfhic

        // Map nfc to NumberFormat 
        level.NumberFormat = nfc switch
        {
            0 => NumberFormat.Decimal,
            1 => NumberFormat.UpperRoman,
            2 => NumberFormat.LowerRoman,
            3 => NumberFormat.UpperLetter,
            4 => NumberFormat.LowerLetter,
            5 => NumberFormat.OrdinalNumber,
            6 => NumberFormat.CardinalText,
            7 => NumberFormat.OrdinalText,
            10 => NumberFormat.DecimalFullWidth,
            11 => NumberFormat.DecimalHalfWidth,
            12 => NumberFormat.JapaneseCounting,
            14 => NumberFormat.DecimalEnclosedCircle,
            22 => NumberFormat.DecimalZero,
            23 => NumberFormat.Bullet,
            30 => NumberFormat.TaiwaneseCountingThousand,
            31 => NumberFormat.TaiwaneseDigital,
            33 => NumberFormat.ChineseCounting,
            34 => NumberFormat.ChineseCountingThousand,
            35 => NumberFormat.KoreanDigital,
            38 => NumberFormat.KoreanCounting,
            39 => NumberFormat.Hebrew1,
            41 => NumberFormat.ArabicAlpha,
            45 => NumberFormat.Hebrew2,
            46 => NumberFormat.ArabicAbjad,
            47 => NumberFormat.HindiVowels,
            _ => nfc == 23 || (listType == ListType.Bullet && levelIndex == 0) ? NumberFormat.Bullet : NumberFormat.Decimal
        };

        level.Alignment = jcByte & 0x03; // 0=left, 1=center, 2=right

        // Read grpprlPapx (PAP SPRM data)
        if (cbGrpprlPapx > 0 && _tableReader.BaseStream.Position + cbGrpprlPapx <= endPos)
        {
            var papxData = _tableReader.ReadBytes(cbGrpprlPapx);
            // Extract indent from PAP SPRMs (sprmPDxaLeft = 0x840F, sprmPDxaLeft1 = 0x8411)
            level.Indent = ExtractIndentFromSprm(papxData);
        }

        // Read grpprlChpx (CHP SPRM data)
        if (cbGrpprlChpx > 0 && _tableReader.BaseStream.Position + cbGrpprlChpx <= endPos)
        {
            var chpxData = _tableReader.ReadBytes(cbGrpprlChpx);
            // We could parse CHP SPRMs for font/color of the number, but skip for now
            // to avoid complexity. The NumberingWriter handles basic run properties.
        }

        // Read xst (number text string)
        if (_tableReader.BaseStream.Position + 2 <= endPos)
        {
            ushort xstLen = _tableReader.ReadUInt16();
            if (xstLen > 0 && xstLen < 256 && _tableReader.BaseStream.Position + xstLen * 2 <= endPos)
            {
                var xstBytes = _tableReader.ReadBytes(xstLen * 2);
                var xstText = Encoding.Unicode.GetString(xstBytes);

                // Convert placeholder bytes (0x00-0x08 mean level 1-9) to %N format
                var sb = new StringBuilder();
                foreach (char c in xstText)
                {
                    if (c >= '\x00' && c <= '\x08')
                    {
                        sb.Append('%');
                        sb.Append((int)c + 1);
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                level.NumberText = sb.ToString();
            }
        }

        // Defaults
        if (string.IsNullOrEmpty(level.NumberText))
        {
            level.NumberText = level.NumberFormat == NumberFormat.Bullet 
                ? "·" 
                : "%" + (levelIndex + 1);
        }
        if (level.Indent == 0)
        {
            level.Indent = 720 * (levelIndex + 1);
        }

        return level;
    }

    private static ListLevel CreateDefaultLevel(int lvl, ListType listType)
    {
        return new ListLevel
        {
            Level = lvl,
            NumberFormat = lvl == 0 && listType == ListType.Bullet ? NumberFormat.Bullet : NumberFormat.Decimal,
            StartAt = 1,
            Indent = 720 * (lvl + 1),
            NumberText = lvl == 0 && listType == ListType.Bullet ? "·" : "%" + (lvl + 1)
        };
    }

    /// <summary>
    /// Extracts the left indent value from PAP SPRM data.
    /// </summary>
    private static int ExtractIndentFromSprm(byte[] sprmData)
    {
        int pos = 0;
        while (pos + 3 < sprmData.Length)
        {
            ushort sprm = (ushort)(sprmData[pos] | (sprmData[pos + 1] << 8));
            pos += 2;
            int sprmSize = (sprm >> 13) & 0x07;
            int operandSize = sprmSize switch
            {
                0 => 1, // toggle
                1 => 1,
                2 => 2,
                3 => 4,
                4 => 2,
                5 => 2,
                _ => 1
            };
            // sprmPDxaLeft (0x840F) and sprmPDxaLeft1 (0x8411)  
            if (sprm == 0x840F && pos + 2 <= sprmData.Length)
            {
                return (short)(sprmData[pos] | (sprmData[pos + 1] << 8));
            }
            pos += operandSize;
        }
        return 0;
    }

    private void ReadPlfLfo()
    {
        if (_fib.FcPlfLfo == 0 || _fib.LcbPlfLfo == 0)
            return;

        _tableReader.BaseStream.Seek(_fib.FcPlfLfo, SeekOrigin.Begin);
        var endPos = _fib.FcPlfLfo + _fib.LcbPlfLfo;

        if (endPos - _tableReader.BaseStream.Position < 4)
            return;

        var lfoCount = _tableReader.ReadInt32();
        if (lfoCount <= 0 || lfoCount > 1000)
            return;

        for (int i = 0; i < lfoCount && _tableReader.BaseStream.Position + 20 <= endPos; i++)
        {
            try
            {
                var lsbfr = _tableReader.ReadInt32();
                var reserved = _tableReader.ReadInt32();

                var flags = _tableReader.ReadUInt16();
                _tableReader.ReadUInt16();

                var grpLfo = new byte[12];
                for (int j = 0; j < 12 && _tableReader.BaseStream.Position < endPos; j++)
                {
                    grpLfo[j] = _tableReader.ReadByte();
                }

                if (i < ListFormats.Count)
                {
                    ApplyLfoOverrides(ListFormats[i], grpLfo);
                }
            }
            catch
            {
                break;
            }
        }
    }

    private void ApplyLfoOverrides(ListFormat listFormat, byte[] grpLfo)
    {
        if (grpLfo.Length < 12)
            return;

        for (int lvl = 0; lvl < Math.Min(listFormat.Levels.Count, 9); lvl++)
        {
            var offset = lvl * 2;
            if (offset + 1 < grpLfo.Length)
            {
                var startAt = grpLfo[offset] | (grpLfo[offset + 1] << 8);
                if (startAt > 0)
                {
                    listFormat.Levels[lvl].StartAt = startAt;
                }
            }
        }
    }

    public ListFormat? GetListFormat(int listId)
    {
        return ListFormats.FirstOrDefault(l => l.ListId == listId);
    }

    public static bool IsListParagraph(ParagraphModel paragraph)
    {
        if (paragraph.Properties == null) return false;
        return paragraph.ListFormatId > 0 || paragraph.ListLevel > 0;
    }

    public int GetListLevel(ParagraphModel paragraph)
    {
        return paragraph.ListLevel;
    }
}
