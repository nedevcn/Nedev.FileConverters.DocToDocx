using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

public class ListReader
{
    private static readonly (string Text, string FontName)[] DefaultBulletLevelSequence =
    {
        ("\uF0B7", "Symbol"),
        ("o", "Courier New"),
        ("\uF0A7", "Wingdings")
    };

    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;

    public StyleSheet? Styles { get; set; }
    public List<NumberingDefinition> NumberingDefinitions { get; private set; } = new();
    public List<ListFormat> ListFormats { get; private set; } = new();
    public List<ListFormatOverride> ListFormatOverrides { get; private set; } = new();

    public ListReader(BinaryReader tableReader, FibReader fib)
    {
        _tableReader = tableReader;
        _fib = fib;
    }

    public void Read()
    {
        NumberingDefinitions = new List<NumberingDefinition>();
        ListFormats = new List<ListFormat>();
        ListFormatOverrides = new List<ListFormatOverride>();

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
            Logger.Warning("Failed to read list formats", ex);
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

        if (endPos - _tableReader.BaseStream.Position < 2)
            return;

        var lstfCount = _tableReader.ReadUInt16();
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

            ApplyCharacterStyleFallbacks(listFormat);
            ApplyBuiltInBulletFallbacks(listFormat);

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
            var chp = new ChpBase();
            var sprmParser = new SprmParser(_tableReader, 0);
            sprmParser.ApplyToChp(chpxData, chp);
            level.RunProperties = ConvertToRunProperties(chp);
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
            NumberFormat = listType == ListType.Bullet ? NumberFormat.Bullet : NumberFormat.Decimal,
            StartAt = 1,
            Indent = 720 * (lvl + 1),
            NumberText = listType == ListType.Bullet ? "·" : "%" + (lvl + 1)
        };
    }

    private void ApplyCharacterStyleFallbacks(ListFormat listFormat)
    {
        if (Styles?.Styles == null || Styles.Styles.Count == 0)
            return;

        foreach (var level in listFormat.Levels)
        {
            if (!string.IsNullOrWhiteSpace(level.RunProperties?.FontName))
                continue;

            var style = Styles.Styles.FirstOrDefault(candidate =>
                candidate.Type == StyleType.Character &&
                string.Equals(candidate.Name, $"WW8Num{listFormat.ListId}z{level.Level}", StringComparison.OrdinalIgnoreCase));

            style ??= Styles.Styles.FirstOrDefault(candidate =>
                candidate.Type == StyleType.Character &&
                string.Equals(candidate.Name, $"WW8Num{listFormat.ListId}z0", StringComparison.OrdinalIgnoreCase));

            if (style?.RunProperties == null)
                continue;

            level.RunProperties = CloneRunProperties(style.RunProperties);
        }
    }

    private static void ApplyBuiltInBulletFallbacks(ListFormat listFormat)
    {
        if (listFormat.Type != ListType.Bullet || listFormat.Levels.Count == 0)
            return;

        bool hasExplicitGlyphs = listFormat.Levels.Any(level =>
            !string.IsNullOrWhiteSpace(level.NumberText) &&
            !string.Equals(level.NumberText, "\u00B7", StringComparison.Ordinal));

        if (hasExplicitGlyphs)
            return;

        for (int index = 0; index < listFormat.Levels.Count; index++)
        {
            var level = listFormat.Levels[index];
            if (!string.IsNullOrWhiteSpace(level.NumberText) && !string.Equals(level.NumberText, "\u00B7", StringComparison.Ordinal))
                continue;

            var fallback = DefaultBulletLevelSequence[index % DefaultBulletLevelSequence.Length];
            level.NumberText = fallback.Text;
            level.RunProperties ??= new RunProperties();
            level.RunProperties.FontIndex = -1;
            level.RunProperties.FontName = fallback.FontName;
        }
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

        var overrides = new List<ListFormatOverride>();

        for (int i = 0; i < lfoCount && _tableReader.BaseStream.Position + 20 <= endPos; i++)
        {
            try
            {
                var lsbfr = _tableReader.ReadInt32();
                _tableReader.ReadInt32();

                _tableReader.ReadUInt16();
                _tableReader.ReadUInt16();

                var grpLfo = new byte[12];
                for (int j = 0; j < 12 && _tableReader.BaseStream.Position < endPos; j++)
                {
                    grpLfo[j] = _tableReader.ReadByte();
                }

                var listId = lsbfr;
                if (listId <= 0 || !ListFormats.Any(format => format.ListId == listId))
                {
                    listId = i < ListFormats.Count ? ListFormats[i].ListId : 0;
                }

                if (listId <= 0)
                {
                    continue;
                }

                overrides.Add(CreateListFormatOverride(i + 1, listId, grpLfo));
            }
            catch
            {
                break;
            }
        }

        ListFormatOverrides = NormalizeListFormatOverrides(overrides);
    }

    private List<ListFormatOverride> NormalizeListFormatOverrides(List<ListFormatOverride> overrides)
    {
        var normalizedOverrides = overrides
            .Where(overrideDefinition => overrideDefinition.OverrideId > 0)
            .GroupBy(overrideDefinition => overrideDefinition.OverrideId)
            .Select(group => group.First())
            .OrderBy(overrideDefinition => overrideDefinition.OverrideId)
            .ToList();

        var existingIds = normalizedOverrides
            .Select(overrideDefinition => overrideDefinition.OverrideId)
            .ToHashSet();

        foreach (var listFormat in ListFormats.Where(format => format.ListId > 0))
        {
            if (existingIds.Contains(listFormat.ListId))
            {
                continue;
            }

            normalizedOverrides.Add(new ListFormatOverride
            {
                OverrideId = listFormat.ListId,
                ListId = listFormat.ListId
            });
        }

        normalizedOverrides.Sort((left, right) => left.OverrideId.CompareTo(right.OverrideId));
        return normalizedOverrides;
    }

    private static ListFormatOverride CreateListFormatOverride(int overrideId, int listId, byte[] grpLfo)
    {
        var listOverride = new ListFormatOverride
        {
            OverrideId = overrideId,
            ListId = listId
        };

        for (int lvl = 0; lvl < Math.Min(grpLfo.Length / 2, 9); lvl++)
        {
            var offset = lvl * 2;
            if (offset + 1 >= grpLfo.Length)
            {
                continue;
            }

            var startAt = grpLfo[offset] | (grpLfo[offset + 1] << 8);
            if (startAt <= 0)
            {
                continue;
            }

            listOverride.Levels.Add(new ListLevelOverride
            {
                Level = lvl,
                StartAt = startAt
            });
        }

        return listOverride;
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

    private RunProperties ConvertToRunProperties(ChpBase chp)
    {
        RunProperties? baseProps = null;
        if (chp.StyleId != 0)
        {
            baseProps = GetRunPropertiesFromCharacterStyle(chp.StyleId);
        }

        var props = new RunProperties
        {
            FontIndex = chp.FontIndex,
            FontSize = chp.FontSize,
            FontSizeCs = chp.FontSizeCs,
            IsBold = chp.IsBold,
            IsBoldCs = chp.IsBoldCs,
            IsItalic = chp.IsItalic,
            IsItalicCs = chp.IsItalicCs,
            IsUnderline = chp.IsUnderline,
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
            IsDeleted = chp.IsDeleted,
            IsInserted = chp.IsInserted,
            AuthorIndexDel = chp.AuthorIndexDel,
            AuthorIndexIns = chp.AuthorIndexIns,
            DateDel = chp.DateDel,
            DateIns = chp.DateIns
        };
        props.FontName = ResolveFontName(chp.FontIndex);

        return baseProps == null ? props : MergeRunProperties(baseProps, props);
    }

    private RunProperties? GetRunPropertiesFromCharacterStyle(ushort styleId)
    {
        if (Styles?.Styles == null || Styles.Styles.Count == 0)
            return null;

        var style = Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Character && s.StyleId == styleId);
        return style?.RunProperties == null ? null : CloneRunProperties(style.RunProperties);
    }

    private string? ResolveFontName(int fontIndex)
    {
        if (Styles == null || fontIndex < 0 || fontIndex >= Styles.Fonts.Count)
            return null;

        var font = Styles.Fonts[fontIndex];
        if (!string.IsNullOrWhiteSpace(font.Name) &&
            !string.Equals(font.Name, $"Font{font.Index}", StringComparison.OrdinalIgnoreCase))
        {
            return font.Name;
        }

        return string.IsNullOrWhiteSpace(font.AltName) ? null : font.AltName;
    }

    private static RunProperties CloneRunProperties(RunProperties source)
    {
        return new RunProperties
        {
            FontIndex = source.FontIndex,
            FontName = source.FontName,
            FontSize = source.FontSize,
            FontSizeCs = source.FontSizeCs,
            IsBold = source.IsBold,
            IsBoldCs = source.IsBoldCs,
            IsItalic = source.IsItalic,
            IsItalicCs = source.IsItalicCs,
            IsUnderline = source.IsUnderline,
            UnderlineType = source.UnderlineType,
            IsStrikeThrough = source.IsStrikeThrough,
            IsDoubleStrikeThrough = source.IsDoubleStrikeThrough,
            IsSmallCaps = source.IsSmallCaps,
            IsAllCaps = source.IsAllCaps,
            IsHidden = source.IsHidden,
            IsSuperscript = source.IsSuperscript,
            IsSubscript = source.IsSubscript,
            Color = source.Color,
            BgColor = source.BgColor,
            CharacterSpacingAdjustment = source.CharacterSpacingAdjustment,
            Language = source.Language,
            LanguageAsia = source.LanguageAsia,
            LanguageCs = source.LanguageCs,
            HighlightColor = source.HighlightColor,
            RgbColor = source.RgbColor,
            HasRgbColor = source.HasRgbColor,
            IsOutline = source.IsOutline,
            IsShadow = source.IsShadow,
            IsEmboss = source.IsEmboss,
            IsImprint = source.IsImprint,
            Border = source.Border,
            Kerning = source.Kerning,
            Position = source.Position,
            CharacterScale = source.CharacterScale,
            EastAsianLayoutType = source.EastAsianLayoutType,
            IsEastAsianVertical = source.IsEastAsianVertical,
            IsEastAsianVerticalCompress = source.IsEastAsianVerticalCompress,
            SnapToGrid = source.SnapToGrid,
            RubyText = source.RubyText,
            IsDeleted = source.IsDeleted,
            IsInserted = source.IsInserted,
            AuthorIndexDel = source.AuthorIndexDel,
            AuthorIndexIns = source.AuthorIndexIns,
            DateDel = source.DateDel,
            DateIns = source.DateIns
        };
    }

    private static RunProperties MergeRunProperties(RunProperties baseProps, RunProperties directProps)
    {
        var merged = CloneRunProperties(baseProps);

        if (directProps.FontIndex != -1)
            merged.FontIndex = directProps.FontIndex;
        if (!string.IsNullOrEmpty(directProps.FontName))
            merged.FontName = directProps.FontName;
        if (directProps.FontSize != 24)
            merged.FontSize = directProps.FontSize;
        if (directProps.FontSizeCs != 24)
            merged.FontSizeCs = directProps.FontSizeCs;

        merged.IsBold = directProps.IsBold || merged.IsBold;
        merged.IsBoldCs = directProps.IsBoldCs || merged.IsBoldCs;
        merged.IsItalic = directProps.IsItalic || merged.IsItalic;
        merged.IsItalicCs = directProps.IsItalicCs || merged.IsItalicCs;
        merged.IsUnderline = directProps.IsUnderline || merged.IsUnderline;
        if (directProps.UnderlineType != UnderlineType.None)
            merged.UnderlineType = directProps.UnderlineType;
        merged.IsStrikeThrough = directProps.IsStrikeThrough || merged.IsStrikeThrough;
        merged.IsDoubleStrikeThrough = directProps.IsDoubleStrikeThrough || merged.IsDoubleStrikeThrough;
        merged.IsSmallCaps = directProps.IsSmallCaps || merged.IsSmallCaps;
        merged.IsAllCaps = directProps.IsAllCaps || merged.IsAllCaps;
        merged.IsHidden = directProps.IsHidden || merged.IsHidden;
        merged.IsSuperscript = directProps.IsSuperscript || merged.IsSuperscript;
        merged.IsSubscript = directProps.IsSubscript || merged.IsSubscript;
        merged.IsOutline = directProps.IsOutline || merged.IsOutline;
        merged.IsShadow = directProps.IsShadow || merged.IsShadow;
        merged.IsEmboss = directProps.IsEmboss || merged.IsEmboss;
        merged.IsImprint = directProps.IsImprint || merged.IsImprint;
        if (directProps.Border != null)
            merged.Border = directProps.Border;
        if (directProps.HasRgbColor)
        {
            merged.RgbColor = directProps.RgbColor;
            merged.HasRgbColor = true;
            merged.Color = directProps.Color;
        }
        else if (directProps.Color != 0)
        {
            merged.Color = directProps.Color;
        }

        if (directProps.BgColor != -1)
            merged.BgColor = directProps.BgColor;
        if (directProps.HighlightColor != 0)
            merged.HighlightColor = directProps.HighlightColor;
        if (directProps.CharacterSpacingAdjustment != 0)
            merged.CharacterSpacingAdjustment = directProps.CharacterSpacingAdjustment;
        if (directProps.Kerning != 0)
            merged.Kerning = directProps.Kerning;
        if (directProps.Position != 0)
            merged.Position = directProps.Position;
        if (directProps.CharacterScale != 100)
            merged.CharacterScale = directProps.CharacterScale;
        if (directProps.EastAsianLayoutType != 0)
            merged.EastAsianLayoutType = directProps.EastAsianLayoutType;
        merged.IsEastAsianVertical = directProps.IsEastAsianVertical || merged.IsEastAsianVertical;
        merged.IsEastAsianVerticalCompress = directProps.IsEastAsianVerticalCompress || merged.IsEastAsianVerticalCompress;
        if (!directProps.SnapToGrid)
            merged.SnapToGrid = false;
        if (directProps.Language != 0)
            merged.Language = directProps.Language;
        if (!string.IsNullOrEmpty(directProps.LanguageAsia))
            merged.LanguageAsia = directProps.LanguageAsia;
        if (!string.IsNullOrEmpty(directProps.LanguageCs))
            merged.LanguageCs = directProps.LanguageCs;
        if (!string.IsNullOrEmpty(directProps.RubyText))
            merged.RubyText = directProps.RubyText;
        merged.IsDeleted = directProps.IsDeleted || merged.IsDeleted;
        merged.IsInserted = directProps.IsInserted || merged.IsInserted;
        if (directProps.AuthorIndexDel != 0)
            merged.AuthorIndexDel = directProps.AuthorIndexDel;
        if (directProps.AuthorIndexIns != 0)
            merged.AuthorIndexIns = directProps.AuthorIndexIns;
        if (directProps.DateDel != 0)
            merged.DateDel = directProps.DateDel;
        if (directProps.DateIns != 0)
            merged.DateIns = directProps.DateIns;

        return merged;
    }
}
