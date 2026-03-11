#nullable enable
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.Utils;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class SprmParserTests
{
    [Fact]
    public void ReadChpProperties_PreservesFormattingBeyondMainStoryRange()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var fib = CreateSyntheticFibReader(fib =>
        {
            SetAutoProperty(fib, nameof(FibReader.FcMin), (uint)100);
            SetAutoProperty(fib, nameof(FibReader.CcpText), 5);
            SetAutoProperty(fib, nameof(FibReader.CcpFtn), 10);
            SetAutoProperty(fib, nameof(FibReader.FcPlcfBteChpx), (uint)0);
            SetAutoProperty(fib, nameof(FibReader.LcbPlcfBteChpx), (uint)12);
        });

        using var wordStream = new MemoryStream(BuildChpxWordStreamPage(100, 110, 115));
        using var tableStream = new MemoryStream(BuildBteTable(1));
        using var wordReader = new BinaryReader(wordStream);
        using var tableReader = new BinaryReader(tableStream);

        var textReader = new Nedev.FileConverters.DocToDocx.Readers.TextReader(wordReader, tableReader, fib);
        SeedPieces(textReader, startCp: 10, endCp: 15, fileOffset: 110);
        var parser = new FkpParser(wordReader, tableReader, fib, textReader);

        var properties = parser.ReadChpProperties();

        Assert.Contains(10, properties.Keys);
        Assert.Contains(14, properties.Keys);
        Assert.DoesNotContain(15, properties.Keys);
    }

    [Fact]
    public void ReadPapProperties_PreservesFormattingBeyondMainStoryRange()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var fib = CreateSyntheticFibReader(fib =>
        {
            SetAutoProperty(fib, nameof(FibReader.FcMin), (uint)100);
            SetAutoProperty(fib, nameof(FibReader.CcpText), 5);
            SetAutoProperty(fib, nameof(FibReader.CcpFtn), 10);
            SetAutoProperty(fib, nameof(FibReader.FcPlcfBtePapx), (uint)4);
            SetAutoProperty(fib, nameof(FibReader.LcbPlcfBtePapx), (uint)16);
        });

        using var wordStream = new MemoryStream(BuildPapxWordStreamPage(100, 110, 115));
        using var tableStream = new MemoryStream(BuildBteTable(1, offset: 4));
        using var wordReader = new BinaryReader(wordStream);
        using var tableReader = new BinaryReader(tableStream);

        var textReader = new Nedev.FileConverters.DocToDocx.Readers.TextReader(wordReader, tableReader, fib);
        SeedPieces(textReader, startCp: 10, endCp: 15, fileOffset: 110);
        var parser = new FkpParser(wordReader, tableReader, fib, textReader);

        var properties = parser.ReadPapProperties();

        Assert.Contains(10, properties.Keys);
        Assert.Contains(14, properties.Keys);
        Assert.DoesNotContain(15, properties.Keys);
    }

    [Fact]
    public void ApplyToChp_DecodesWord97CharacterOpcodesByFullCode()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var chp = new ChpBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyChpSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x0835, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x0836, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4A43, 44);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4852, 200);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x484B, 16);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x0854, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x085C, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x085D, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4863, 7);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6864, 0x12345678);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4873, 0x0409);

        Assert.True(chp.IsBold);
        Assert.True(chp.IsItalic);
        Assert.Equal(44, chp.FontSize);
        Assert.Equal(200, chp.Scale);
        Assert.Equal(16, chp.Kerning);
        Assert.True(chp.IsImprint);
        Assert.True(chp.IsBoldCs);
        Assert.True(chp.IsItalicCs);
        Assert.Equal((ushort)7, chp.AuthorIndexDel);
        Assert.Equal(0x12345678u, chp.DateDel);
        Assert.Equal(0x0409, chp.LanguageId);
    }

    [Fact]
    public void ApplyToChp_RsidSprms_DoNotTriggerWord6ShadowOrEmbossFallbacks()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var chp = new ChpBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyChpSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6815, 0x01020304);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6816, 0x01020304);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6817, 0x01020304);

        Assert.False(chp.IsShadow);
        Assert.False(chp.IsEmboss);
    }

    [Fact]
    public void ApplyToPap_DecodesCharacterUnitIndents()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var pap = new PapBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyPapSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        ApplySprm(parser, sprmType!, applyMethod!, pap, WordConsts.SPRM_PCHTO, 206);
        ApplySprm(parser, sprmType!, applyMethod!, pap, WordConsts.SPRM_PCHTO2, 32);
        ApplySprm(parser, sprmType!, applyMethod!, pap, WordConsts.SPRM_PCHTO3, 18);

        Assert.Equal(206, pap.IndentFirstLineChars);
        Assert.Equal(32, pap.IndentLeftChars);
        Assert.Equal(18, pap.IndentRightChars);
    }

    [Fact]
    public void ApplyToPap_DecodesLegacyParagraphSpacingAndJustificationSprms()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var pap = new PapBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyPapSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        ApplySprm(parser, sprmType!, applyMethod!, pap, 0x2403, 2);
        ApplySprm(parser, sprmType!, applyMethod!, pap, 0x4459, 51);
        ApplySprm(parser, sprmType!, applyMethod!, pap, 0x6412, 0x00010114);

        Assert.Equal(2, pap.Justification);
        Assert.Equal(276, pap.LineSpacing);
        Assert.Equal(1, pap.LineSpacingMultiple);
        Assert.Equal(51, pap.SpaceAfterLines);
        Assert.Equal(0, pap.SpaceAfter);
    }

    [Fact]
    public void ApplyToChp_MetadataSprms_DoNotTriggerWord6FallbackFormatting()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var chp = new ChpBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyChpSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        var metadataSprms = new (ushort code, uint operand)[]
        {
            (0x0802, 1),
            (0x0806, 1),
            (0x080A, 1),
            (0x0811, 1),
            (0x0818, 1),
            (0x0855, 1),
            (0x0856, 1),
            (0x085A, 1),
            (0x0875, 1),
            (0x0882, 1),
            (0x4807, 1),
            (0x4867, 1),
            (0x6A09, 1),
            (0xC81A, 1),
            (0xCA57, 1),
            (0xCA62, 1),
            (0xCA89, 1)
        };

        foreach (var (code, operand) in metadataSprms)
            ApplySprm(parser, sprmType!, applyMethod!, chp, code, operand);

        Assert.False(chp.IsBold);
        Assert.False(chp.IsItalic);
        Assert.False(chp.IsOutline);
        Assert.False(chp.IsShadow);
        Assert.False(chp.IsEmboss);
        Assert.False(chp.IsImprint);
        Assert.False(chp.IsHidden);
        Assert.Equal(-1, chp.FontIndex);
        Assert.Equal(0, chp.Color);
    }

    [Fact]
    public void SampleTextDoc_ScalingRun_ComesFromFkpCharScaleSprm()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)GetPrivateField(docReader, "_textReader")!;
        var globalChpMap = (Dictionary<int, ChpBase>)GetPrivateField(docReader, "_globalChpMap")!;
        var fullText = textReader.Text;
        var marker = "Scaling 200%";
        var markerCp = fullText.IndexOf(marker, StringComparison.Ordinal);

        Assert.True(markerCp >= 0, $"Could not find '{marker}' in sample text. Text excerpt: {TakeExcerpt(fullText, 0, 300)}");

        var scaleCp = markerCp + "Scaling ".Length;
        var piece = textReader.Pieces.FirstOrDefault(p => scaleCp >= p.CpStart && scaleCp < p.CpEnd);
        var pieceChp = textReader.GetPieceRunPropertiesAtCp(scaleCp);
        var directChp = globalChpMap.TryGetValue(scaleCp, out var direct) ? direct : null;
        var pieceModifiers = (Dictionary<ushort, byte[]>)GetPrivateField(textReader, "_piecePropertyModifiers")!;
        var pieceGrpprlHex = ResolvePieceGrpprlHex(piece, pieceModifiers);
        var fkpDetails = GetFkpEntriesForCp(docReader, scaleCp);

        var details = new StringBuilder();
        details.AppendLine($"scaleCp={scaleCp}");
        details.AppendLine($"markerCp={markerCp}");
        details.AppendLine($"piece={FormatPiece(piece)}");
        details.AppendLine($"pieceChp={FormatChp(pieceChp)}");
        details.AppendLine($"directChp={FormatChp(directChp)}");
        details.AppendLine($"pieceGrpprl={pieceGrpprlHex}");
        details.AppendLine("fkpEntries:");
        foreach (var line in fkpDetails)
            details.AppendLine(line);

        Assert.True(piece == null || piece.Prm == 0, details.ToString());
        Assert.True(pieceChp == null || pieceChp.Scale == 100, details.ToString());
        Assert.NotNull(directChp);
        Assert.False(directChp!.IsBold, details.ToString());
        Assert.False(directChp.IsItalic, details.ToString());
        Assert.Equal(200, directChp.Scale);
        Assert.DoesNotContain("52 48", pieceGrpprlHex, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(fkpDetails, line => line.Contains("52 48 C8 00", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void SampleTextDoc_LeadingRuns_ShowWhereShadowComesFrom()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)GetPrivateField(docReader, "_textReader")!;
        var globalChpMap = (Dictionary<int, ChpBase>)GetPrivateField(docReader, "_globalChpMap")!;
        var checkpoints = new[]
        {
            0,
            textReader.Text.IndexOf("居中", StringComparison.Ordinal),
            textReader.Text.IndexOf("粗体", StringComparison.Ordinal),
            textReader.Text.IndexOf("文字Scaling 200%", StringComparison.Ordinal)
        };

        var report = new StringBuilder();
        foreach (var cp in checkpoints.Where(cp => cp >= 0).Distinct())
        {
            var directChp = globalChpMap.TryGetValue(cp, out var direct) ? direct : null;
            report.AppendLine($"cp={cp} direct={FormatChp(directChp)}");
            foreach (var line in GetFkpEntriesForCp(docReader, cp))
                report.AppendLine(line);
        }

        Assert.DoesNotContain("39 08 81", report.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SampleTextDoc_NormalStyle_DoesNotCarryShadow()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var normalStyle = docReader.Document.Styles.Styles.FirstOrDefault(s => s.StyleId == 0 || string.Equals(s.Name, "Normal", StringComparison.OrdinalIgnoreCase));

        Assert.NotNull(normalStyle);
        Assert.False(normalStyle!.RunProperties?.IsShadow ?? false);
    }

    [Fact]
    public void SampleTextDoc_PapMap_CapturesAlignmentAndIndentMetadata()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)GetPrivateField(docReader, "_textReader")!;
        var papMap = (Dictionary<int, PapBase>)GetPrivateField(docReader, "_globalPapMap")!;
        var fib = (FibReader)GetPrivateField(docReader, "_fibReader")!;
        var tableReader = (BinaryReader)GetPrivateField(docReader, "_tableReader")!;
        var wordDocReader = (BinaryReader)GetPrivateField(docReader, "_wordDocReader")!;
        var fkpParser = GetPrivateField(docReader, "_fkpParser")!;
        var papCache = (IDictionary)GetPrivateField(fkpParser, "_papFkpCache")!;
        var pieceModifiers = (Dictionary<ushort, byte[]>)GetPrivateField(textReader, "_piecePropertyModifiers")!;

        var centeredCp = textReader.Text.IndexOf("居中", StringComparison.Ordinal);
        var rightCp = textReader.Text.IndexOf("右对齐", StringComparison.Ordinal);
        var indentCp = textReader.Text.IndexOf("Indent", StringComparison.Ordinal);

        Assert.True(centeredCp >= 0 && rightCp >= 0 && indentCp >= 0, "Failed to locate expected sample markers in reconstructed text.");

        var centeredPap = ResolvePapAtCp(papMap, centeredCp, 4096);
        var rightPap = ResolvePapAtCp(papMap, rightCp, 4096);
        var indentPap = ResolvePapAtCp(papMap, indentCp, 4096);

        var report = new StringBuilder();
        report.AppendLine($"fComplex={fib.FComplex} fcMin=0x{fib.FcMin:X} fcMac=0x{fib.FcMac:X} pieceCount={textReader.Pieces.Count}");
        report.AppendLine($"fcPlcfBtePapx=0x{fib.FcPlcfBtePapx:X} lcbPlcfBtePapx=0x{fib.LcbPlcfBtePapx:X}");
        report.AppendLine($"fcPlcfBteChpx=0x{fib.FcPlcfBteChpx:X} lcbPlcfBteChpx=0x{fib.LcbPlcfBteChpx:X}");
        report.AppendLine($"papBte={DescribeBtePages(tableReader, wordDocReader, fib.FcPlcfBtePapx, fib.LcbPlcfBtePapx)}");
        report.AppendLine($"papFkpCache={DescribePapCache(papCache)}");
        foreach (var line in GetPapFkpEntriesForCp(docReader, rightCp))
            report.AppendLine($"rightPapFkp={line}");
        report.AppendLine($"centerPiecePap={DescribePiecePap(textReader, pieceModifiers, wordDocReader, centeredCp)}");
        report.AppendLine($"rightPiecePap={DescribePiecePap(textReader, pieceModifiers, wordDocReader, rightCp)}");
        report.AppendLine($"indentPiecePap={DescribePiecePap(textReader, pieceModifiers, wordDocReader, indentCp)}");
        report.AppendLine($"papCount={papMap.Count}");
        report.AppendLine($"papKeysAroundCenter={string.Join(",", papMap.Keys.Where(key => Math.Abs(key - centeredCp) <= 8192).OrderBy(key => key).Take(20))}");
        report.AppendLine($"centeredCp={centeredCp} pap={FormatPap(centeredPap)}");
        report.AppendLine($"rightCp={rightCp} pap={FormatPap(rightPap)}");
        report.AppendLine($"indentCp={indentCp} pap={FormatPap(indentPap)}");

        Assert.True(centeredPap != null, report.ToString());
        Assert.True(centeredPap!.Justification == 1, report.ToString());
        Assert.True(rightPap != null, report.ToString());
        Assert.True(rightPap!.Justification == 2, report.ToString());
        Assert.True(indentPap != null, report.ToString());
        Assert.True(indentPap!.IndentLeft > 0, report.ToString());
    }

    [Fact]
    public void ApplyToChp_DecodesCharacterBorderAndEastAsianLayout()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var chp = new ChpBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyChpSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6865, 0x00010108);
        ApplyVariableSprm(parser, sprmType!, applyMethod!, chp, 0xCA78, 0x01, 0x01, 0x10);

        Assert.NotNull(chp.Border);
        Assert.Equal(Nedev.FileConverters.DocToDocx.Models.BorderStyle.Single, chp.Border!.Style);
        Assert.True(chp.IsEastAsianVertical);
        Assert.True(chp.IsEastAsianVerticalCompress);
    }

    [Fact]
    public void SampleTextDoc_PreservesBorderAndVerticalMarkersInChpMap()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)GetPrivateField(docReader, "_textReader")!;
        var globalChpMap = (Dictionary<int, ChpBase>)GetPrivateField(docReader, "_globalChpMap")!;
        var globalPapMap = (Dictionary<int, PapBase>)GetPrivateField(docReader, "_globalPapMap")!;

        var borderCp = textReader.Text.IndexOf("文字边框", StringComparison.Ordinal);
        var verticalCp = textReader.Text.IndexOf("纵向", StringComparison.Ordinal);
        var report = new StringBuilder();
        report.AppendLine($"shapes={docReader.Document.Shapes.Count} textboxes={docReader.Document.Textboxes.Count}");

        Assert.True(borderCp >= 0, "Failed to locate border marker in sample text.");
        Assert.True(verticalCp >= 0, "Failed to locate vertical marker in sample text.");

        var borderChp = ResolveChpAtCp(globalChpMap, borderCp, 32);
        var verticalChp = ResolveChpAtCp(globalChpMap, verticalCp, 32);
        var borderPap = ResolvePapAtCp(globalPapMap, borderCp, 64);
        var verticalPap = ResolvePapAtCp(globalPapMap, verticalCp, 64);

        report.AppendLine($"borderCp={borderCp} chp={FormatChp(borderChp)} pap={FormatPap(borderPap)}");
        foreach (var line in GetFkpEntriesForCp(docReader, borderCp))
            report.AppendLine("  " + line);

        report.AppendLine($"verticalCp={verticalCp} chp={FormatChp(verticalChp)} pap={FormatPap(verticalPap)}");
        foreach (var line in GetFkpEntriesForCp(docReader, verticalCp))
            report.AppendLine("  " + line);

        Assert.NotNull(borderChp);
        Assert.NotNull(borderChp!.Border);
        Assert.True(borderChp.Border!.Style != Nedev.FileConverters.DocToDocx.Models.BorderStyle.None, report.ToString());
        Assert.NotNull(verticalChp);
        Assert.True(verticalChp!.IsEastAsianVertical, report.ToString());
    }

    private static void ApplySprm(SprmParser parser, Type sprmType, MethodInfo applyMethod, ChpBase chp, ushort code, uint operand)
    {
        var sprm = Activator.CreateInstance(sprmType)!;
        sprmType.GetProperty("Code")!.SetValue(sprm, code);
        sprmType.GetProperty("Operand")!.SetValue(sprm, operand);
        sprmType.GetProperty("OperandSize")!.SetValue(sprm, 0);
        applyMethod.Invoke(parser, new object[] { sprm, chp });
    }

    private static void ApplySprm(SprmParser parser, Type sprmType, MethodInfo applyMethod, PapBase pap, ushort code, uint operand)
    {
        var sprm = Activator.CreateInstance(sprmType)!;
        sprmType.GetProperty("Code")!.SetValue(sprm, code);
        sprmType.GetProperty("Operand")!.SetValue(sprm, operand);
        sprmType.GetProperty("OperandSize")!.SetValue(sprm, 0);
        applyMethod.Invoke(parser, new object[] { sprm, pap });
    }

    private static void ApplyVariableSprm(SprmParser parser, Type sprmType, MethodInfo applyMethod, ChpBase chp, ushort code, params byte[] operand)
    {
        var sprm = Activator.CreateInstance(sprmType)!;
        sprmType.GetProperty("Code")!.SetValue(sprm, code);
        sprmType.GetProperty("Operand")!.SetValue(sprm, 0u);
        sprmType.GetProperty("OperandSize")!.SetValue(sprm, 0xFF);
        sprmType.GetProperty("VariableOperand")!.SetValue(sprm, operand);
        applyMethod.Invoke(parser, new object[] { sprm, chp });
    }

    private static object? GetPrivateField(object instance, string fieldName)
    {
        return instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(instance);
    }

    private static FibReader CreateSyntheticFibReader(Action<FibReader> configure)
    {
        var fib = new FibReader(new BinaryReader(new MemoryStream(new byte[512])));
        configure(fib);
        return fib;
    }

    private static void SeedPieces(Nedev.FileConverters.DocToDocx.Readers.TextReader textReader, int startCp, int endCp, uint fileOffset)
    {
        SetPrivateField(textReader, "_pieces", new List<Piece>
        {
            new Piece
            {
                CpStart = startCp,
                CpEnd = endCp,
                FileOffset = fileOffset,
                RawFcMasked = fileOffset * 2,
                IsUnicode = false,
                Prm = 0
            }
        });
    }

    private static void SetAutoProperty<T>(object instance, string propertyName, T value)
    {
        var field = instance.GetType().GetField($"<{propertyName}>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        field!.SetValue(instance, value);
    }

    private static void SetPrivateField<T>(object instance, string fieldName, T value)
    {
        var field = instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        field!.SetValue(instance, value);
    }

    private static byte[] BuildBteTable(uint pn, int offset = 0)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.Default, leaveOpen: true);
        for (var index = 0; index < offset; index++)
            writer.Write((byte)0);
        writer.Write(0);
        writer.Write(5);
        writer.Write(pn);
        if (offset > 0)
            writer.Write(0);
        writer.Flush();
        return stream.ToArray();
    }

    private static byte[] BuildChpxWordStreamPage(int fcMin, int startFc, int endFc)
    {
        var bytes = new byte[WordConsts.FKP_PAGE_SIZE * 2];
        var pageOffset = WordConsts.FKP_PAGE_SIZE;

        BitConverter.GetBytes(startFc).CopyTo(bytes, pageOffset);
        BitConverter.GetBytes(endFc).CopyTo(bytes, pageOffset + 4);
        bytes[pageOffset + 8] = 0;
        bytes[pageOffset + WordConsts.FKP_PAGE_SIZE - 1] = 1;

        return bytes;
    }

    private static byte[] BuildPapxWordStreamPage(int fcMin, int startFc, int endFc)
    {
        var bytes = new byte[WordConsts.FKP_PAGE_SIZE * 2];
        var pageOffset = WordConsts.FKP_PAGE_SIZE;

        BitConverter.GetBytes(startFc).CopyTo(bytes, pageOffset);
        BitConverter.GetBytes(endFc).CopyTo(bytes, pageOffset + 4);
        bytes[pageOffset + 8] = 20;
        bytes[pageOffset + 40] = 0;
        bytes[pageOffset + 41] = 0;
        bytes[pageOffset + WordConsts.FKP_PAGE_SIZE - 1] = 1;

        return bytes;
    }

    private static string ResolvePieceGrpprlHex(Piece? piece, Dictionary<ushort, byte[]> pieceModifiers)
    {
        if (piece == null || piece.Prm == 0)
            return "<none>";

        var candidateKeys = new[]
        {
            piece.Prm,
            (ushort)(piece.Prm & 0xFFFE),
            (ushort)(piece.Prm >> 1),
            (ushort)(piece.Prm & 0x7FFF),
            (ushort)((piece.Prm & 0x7FFF) >> 1)
        };

        foreach (var key in candidateKeys)
        {
            if (pieceModifiers.TryGetValue(key, out var grpprl))
                return $"key=0x{key:X4} bytes={BitConverter.ToString(grpprl).Replace('-', ' ')}";
        }

        return $"unresolved-prm=0x{piece.Prm:X4}";
    }

    private static IReadOnlyList<string> GetFkpEntriesForCp(DocReader docReader, int cp)
    {
        var fkpParser = GetPrivateField(docReader, "_fkpParser")!;
        var cache = (IDictionary)GetPrivateField(fkpParser, "_chpFkpCache")!;
        var lines = new List<string>();

        foreach (DictionaryEntry entry in cache)
        {
            var fkp = entry.Value;
            var entriesProp = fkp!.GetType().GetProperty("Entries")!;
            var entries = (IEnumerable)entriesProp.GetValue(fkp)!;

            foreach (var item in entries)
            {
                var startCp = (int)item!.GetType().GetProperty("StartCpOffset")!.GetValue(item)!;
                var endCp = (int)item.GetType().GetProperty("EndCpOffset")!.GetValue(item)!;
                if (cp < startCp || cp >= endCp)
                    continue;

                var rawGrpprl = (byte[])item.GetType().GetProperty("RawGrpprl")!.GetValue(item)!;
                var chp = (ChpBase)item.GetType().GetProperty("Properties")!.GetValue(item)!;
                lines.Add($"pn={entry.Key} cp={startCp}..{endCp} grpprl={BitConverter.ToString(rawGrpprl).Replace('-', ' ')} chp={FormatChp(chp)}");
            }
        }

        if (lines.Count == 0)
            lines.Add("<no fkp entry for cp>");

        return lines;
    }

    private static IReadOnlyList<string> GetPapFkpEntriesForCp(DocReader docReader, int cp)
    {
        var fkpParser = GetPrivateField(docReader, "_fkpParser")!;
        var cache = (IDictionary)GetPrivateField(fkpParser, "_papFkpCache")!;
        var lines = new List<string>();

        foreach (DictionaryEntry entry in cache)
        {
            var fkp = entry.Value;
            var entriesProp = fkp!.GetType().GetProperty("Entries")!;
            var entries = (IEnumerable)entriesProp.GetValue(fkp)!;

            foreach (var item in entries)
            {
                var startCp = (int)item!.GetType().GetProperty("StartCpOffset")!.GetValue(item)!;
                var endCp = (int)item.GetType().GetProperty("EndCpOffset")!.GetValue(item)!;
                if (cp < startCp || cp >= endCp)
                    continue;

                var rawGrpprl = (byte[])item.GetType().GetProperty("RawGrpprl")!.GetValue(item)!;
                var pap = (PapBase)item.GetType().GetProperty("Properties")!.GetValue(item)!;
                lines.Add($"pn={entry.Key} cp={startCp}..{endCp} grpprl={BitConverter.ToString(rawGrpprl).Replace('-', ' ')} pap={FormatPap(pap)}");
            }
        }

        if (lines.Count == 0)
            lines.Add("<no pap fkp entry for cp>");

        return lines;
    }

    private static string FormatPiece(Piece? piece)
    {
        if (piece == null)
            return "<none>";

        return $"cp={piece.CpStart}..{piece.CpEnd} prm=0x{piece.Prm:X4} offset=0x{piece.FileOffset:X8} raw=0x{piece.RawFcMasked:X8} unicode={piece.IsUnicode}";
    }

    private static string FormatChp(ChpBase? chp)
    {
        if (chp == null)
            return "<none>";

        return $"bold={chp.IsBold} italic={chp.IsItalic} underline={chp.Underline} scale={chp.Scale} kern={chp.Kerning} size={chp.FontSize} color={chp.Color} highlight={chp.HighlightColor} pos={chp.Position} lang={chp.LanguageId} border={chp.Border?.Style}/{chp.Border?.Width} eastAsiaType={chp.EastAsianLayoutType} vert={chp.IsEastAsianVertical} vertCompress={chp.IsEastAsianVerticalCompress}";
    }

    private static ChpBase? ResolveChpAtCp(Dictionary<int, ChpBase> chpMap, int cp, int maxLookaround)
    {
        if (chpMap.TryGetValue(cp, out var chp))
            return chp;

        for (var offset = 1; offset <= maxLookaround; offset++)
        {
            if (chpMap.TryGetValue(cp - offset, out chp))
                return chp;
            if (chpMap.TryGetValue(cp + offset, out chp))
                return chp;
        }

        return null;
    }

    private static string TakeExcerpt(string text, int start, int length)
    {
        if (text.Length == 0)
            return string.Empty;

        var safeStart = Math.Max(0, Math.Min(start, text.Length - 1));
        var safeLength = Math.Min(length, text.Length - safeStart);
        return text.Substring(safeStart, safeLength).Replace("\r", "\\r").Replace("\n", "\\n");
    }

    private static PapBase? ResolvePapAtCp(Dictionary<int, PapBase> papMap, int cp, int maxLookback)
    {
        if (papMap.TryGetValue(cp, out var pap))
            return pap;

        for (var probe = cp - 1; probe >= Math.Max(0, cp - maxLookback); probe--)
        {
            if (papMap.TryGetValue(probe, out pap))
                return pap;
        }

        return null;
    }

    private static string FormatPap(PapBase? pap)
    {
        if (pap == null)
            return "<none>";

        return $"just={pap.Justification} istd={pap.Istd} style={pap.StyleId} left={pap.IndentLeft} leftChars={pap.IndentLeftChars} first={pap.IndentFirstLine} firstChars={pap.IndentFirstLineChars} right={pap.IndentRight} rightChars={pap.IndentRightChars} before={pap.SpaceBefore} after={pap.SpaceAfter} line={pap.LineSpacing} mult={pap.LineSpacingMultiple} list={pap.ListFormatId}/{pap.ListLevel}";
    }

    private static string DescribeBtePages(BinaryReader tableReader, BinaryReader wordDocReader, uint fc, uint lcb)
    {
        if (fc == 0 || lcb < 12)
            return "<none>";

        var originalTablePosition = tableReader.BaseStream.Position;
        var originalWordPosition = wordDocReader.BaseStream.Position;

        try
        {
            tableReader.BaseStream.Seek(fc, SeekOrigin.Begin);
            var entryCount = (int)((lcb - 4) / 8);
            var cps = new int[entryCount + 1];
            for (var index = 0; index <= entryCount; index++)
                cps[index] = tableReader.ReadInt32();

            var descriptions = new List<string>();
            for (var index = 0; index < entryCount; index++)
            {
                var pn = tableReader.ReadUInt32();
                var offset = pn * WordConsts.FKP_PAGE_SIZE;
                byte crun = 0;
                if (offset + WordConsts.FKP_PAGE_SIZE <= wordDocReader.BaseStream.Length)
                {
                    wordDocReader.BaseStream.Seek(offset + WordConsts.FKP_PAGE_SIZE - 1, SeekOrigin.Begin);
                    crun = wordDocReader.ReadByte();
                }

                descriptions.Add($"{index}:cp={cps[index]}..{cps[index + 1]} pn={pn} crun={crun}");
            }

            return string.Join(" | ", descriptions);
        }
        finally
        {
            tableReader.BaseStream.Seek(originalTablePosition, SeekOrigin.Begin);
            wordDocReader.BaseStream.Seek(originalWordPosition, SeekOrigin.Begin);
        }
    }

    private static string DescribePapCache(IDictionary papCache)
    {
        var descriptions = new List<string>();

        foreach (DictionaryEntry entry in papCache)
        {
            var fkp = entry.Value;
            var entriesProp = fkp!.GetType().GetProperty("Entries")!;
            var entries = ((IEnumerable)entriesProp.GetValue(fkp)!).Cast<object>().ToList();
            var ranges = entries
                .Take(4)
                .Select(item => $"{item.GetType().GetProperty("StartCpOffset")!.GetValue(item)}..{item.GetType().GetProperty("EndCpOffset")!.GetValue(item)}")
                .ToList();

            descriptions.Add($"pn={entry.Key} count={entries.Count} ranges=[{string.Join(",", ranges)}]");
        }

        return descriptions.Count == 0 ? "<empty>" : string.Join(" | ", descriptions);
    }

    private static string DescribePiecePap(Nedev.FileConverters.DocToDocx.Readers.TextReader textReader, Dictionary<ushort, byte[]> pieceModifiers, BinaryReader wordDocReader, int cp)
    {
        var piece = textReader.Pieces.FirstOrDefault(candidate => cp >= candidate.CpStart && cp < candidate.CpEnd);
        if (piece == null)
            return "<no piece>";

        var candidateKeys = new[]
        {
            piece.Prm,
            (ushort)(piece.Prm & 0xFFFE),
            (ushort)(piece.Prm >> 1),
            (ushort)(piece.Prm & 0x7FFF),
            (ushort)((piece.Prm & 0x7FFF) >> 1)
        };

        foreach (var key in candidateKeys)
        {
            if (!pieceModifiers.TryGetValue(key, out var grpprl))
                continue;

            var pap = new PapBase();
            new SprmParser(wordDocReader, 0).ApplyToPap(grpprl, pap);
            return $"piece={FormatPiece(piece)} key=0x{key:X4} grpprl={BitConverter.ToString(grpprl).Replace('-', ' ')} pap={FormatPap(pap)}";
        }

        return $"piece={FormatPiece(piece)} no-grpprl";
    }
}