#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Nedev.FileConverters.Core;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class DocToDocxConverterTests
{
    [Fact]
    public void Convert_DetectsDocxBySignature_AndCopiesWithoutLibraryConsoleOutput()
    {
        var inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".bin");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var originalOut = Console.Out;

        try
        {
            var document = new DocumentModel();
            document.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "copied" } } });
            DocToDocxConverter.SaveDocument(document, sourcePath);
            File.Copy(sourcePath, inputPath, overwrite: true);

            using var writer = new StringWriter();
            Console.SetOut(writer);

            DocToDocxConverter.Convert(inputPath, outputPath);

            Assert.True(DocToDocxConverter.ValidatePackage(outputPath, out var validationError), validationError);
            Assert.Equal(File.ReadAllBytes(inputPath), File.ReadAllBytes(outputPath));
            Assert.Equal(string.Empty, writer.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            DeleteIfExists(inputPath);
            DeleteIfExists(outputPath);
            DeleteIfExists(sourcePath);
        }
    }

    [Fact]
    public void Convert_ReportsErrorProgress_WhenInputFormatIsUnsupported()
    {
        var inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".bin");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        File.WriteAllBytes(inputPath, Encoding.UTF8.GetBytes("not-a-word-document"));

        var updates = new List<ConversionProgress>();
        var progress = new ImmediateProgress<ConversionProgress>(updates.Add);

        try
        {
            var exception = Assert.Throws<InvalidDataException>(() =>
                DocToDocxConverter.Convert(inputPath, outputPath, progress, password: null, enableHyperlinks: true));

            Assert.Contains("Unsupported input format", exception.Message, StringComparison.Ordinal);
            Assert.NotEmpty(updates);
            Assert.Equal(ConversionStage.Error, updates[^1].Stage);
            Assert.Contains("Unsupported input format", updates[^1].Message, StringComparison.Ordinal);
        }
        finally
        {
            DeleteIfExists(inputPath);
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void ConvertWithWarnings_DetectsDocxBySignature_AndReturnsNoWarnings()
    {
        var inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".bin");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            var document = new DocumentModel();
            document.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "copied" } } });
            DocToDocxConverter.SaveDocument(document, sourcePath);
            File.Copy(sourcePath, inputPath, overwrite: true);

            var result = DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            Assert.True(DocToDocxConverter.ValidatePackage(outputPath, out var validationError), validationError);
            Assert.Equal(outputPath, result.OutputPath);
            Assert.Empty(result.Warnings);
            Assert.Empty(result.Diagnostics);
        }
        finally
        {
            DeleteIfExists(inputPath);
            DeleteIfExists(outputPath);
            DeleteIfExists(sourcePath);
        }
    }

    [Fact]
    public void Logger_BeginWarningCapture_CapturesWarningsWithinScopeOnly()
    {
        var warnings = new List<string>();

        using (Logger.BeginWarningCapture(warnings))
        {
            Logger.Warning("captured warning");
        }

        Logger.Warning("outside scope");

        Assert.Single(warnings);
        Assert.Contains("captured warning", warnings[0], StringComparison.Ordinal);
        Assert.DoesNotContain(warnings, warning => warning.Contains("outside scope", StringComparison.Ordinal));
    }

    [Fact]
    public void Logger_BeginDiagnosticCapture_CapturesStructuredDiagnosticsWithinScopeOnly()
    {
        var diagnostics = new List<ConversionDiagnostic>();

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            Logger.Warning("captured diagnostic");
        }

        Logger.Warning("outside diagnostic scope");

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal(Logger.LogLevel.Warning, diagnostic.Level);
        Assert.Equal("captured diagnostic", diagnostic.Message);
        Assert.Contains("captured diagnostic", diagnostic.FormattedMessage, StringComparison.Ordinal);
        Assert.Null(diagnostic.ExceptionType);
        Assert.Null(diagnostic.ExceptionMessage);
    }

    [Fact]
    public void ConversionResult_ConstructedFromWarnings_BackfillsDiagnostics()
    {
        var result = new ConversionResult("out.docx", new[] { "warning one" });

        Assert.Single(result.Warnings);
        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(Logger.LogLevel.Warning, diagnostic.Level);
        Assert.Equal("warning one", diagnostic.Message);
        Assert.Equal("warning one", diagnostic.FormattedMessage);
    }

    [Fact]
    public void ValidatePackage_ReturnsFalse_WhenRequiredPartIsMissing()
    {
        var packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            using (var stream = File.Create(packagePath))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" />");
                WriteEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" />");
            }

            Assert.False(DocToDocxConverter.ValidatePackage(packagePath, out var validationError));
            Assert.Contains("word/document.xml", validationError, StringComparison.Ordinal);
        }
        finally
        {
            DeleteIfExists(packagePath);
        }
    }

    [Fact]
    public void ValidatePackage_ReturnsFalse_WhenInternalRelationshipTargetIsMissing()
    {
        var packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            using (var stream = File.Create(packagePath))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteEntry(archive, "[Content_Types].xml", """
                                        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                                            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                                            <Default Extension="xml" ContentType="application/xml" />
                                            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
                    </Types>
                    """);
                WriteEntry(archive, "_rels/.rels", """
                                        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
                    </Relationships>
                    """);
                WriteEntry(archive, "word/document.xml", """
                                        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:body><w:p /></w:body>
                    </w:document>
                    """);
                WriteEntry(archive, "word/_rels/document.xml.rels", """
                                        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                                            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png" />
                    </Relationships>
                    """);
            }

            Assert.False(DocToDocxConverter.ValidatePackage(packagePath, out var validationError));
            Assert.Contains("media/image1.png", validationError, StringComparison.Ordinal);
        }
        finally
        {
            DeleteIfExists(packagePath);
        }
    }

    [Fact]
    public void FileConverter_Convert_BuffersNonSeekableInput()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");
        var inputBytes = File.ReadAllBytes(inputPath);

        using var input = new NonSeekableReadOnlyStream(inputBytes);
        var converter = new DocToDocxFileConverter();
        using var output = converter.Convert(input);

        using var archive = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
        Assert.NotNull(archive.GetEntry("word/document.xml"));
        Assert.True(output.Length > 0);
    }

    [Fact]
    public void Convert_StreamBasedApi_ReturnsValidPackage()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");
        using var input = File.OpenRead(inputPath);
        using var outputStream = new MemoryStream();

        DocToDocxConverter.Convert(input, outputStream, password: null, enableHyperlinks: true);
        outputStream.Position = 0;
        using var archive = new ZipArchive(outputStream, ZipArchiveMode.Read, leaveOpen: true);
        Assert.NotNull(archive.GetEntry("word/document.xml"));
    }

    [Fact]
    public void ConvertWithWarnings_StreamBasedApi_PopulatesResult()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");
        using var input = File.OpenRead(inputPath);
        using var outputStream = new MemoryStream();

        var result = DocToDocxConverter.ConvertWithWarnings(input, outputStream);
        Assert.NotNull(result);
        // when converting using streams there is no file path available, so we
        // simply ensure the property is at least non-null and allow it to be
        // empty.
        Assert.NotNull(result.OutputPath);
        // diagnostics list may be empty when no warnings are generated, but it
        // should never be null.
        Assert.NotNull(result.Diagnostics);
    }

    [Fact]
    public void LoadDocument_Sample1Doc_RecoversRichLayoutContent()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");

        var document = DocToDocxConverter.LoadDocument(inputPath);
        var fullText = string.Concat(document.Paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));

        Assert.True(document.Paragraphs.Count >= 90);
        Assert.True(document.Tables.Count >= 5);
        Assert.True(document.Images.Count >= 2);
        Assert.Contains("Text Formatting", fullText, StringComparison.Ordinal);
        Assert.Contains("Inline formatting", fullText, StringComparison.Ordinal);
        Assert.Contains("Footnotes", fullText, StringComparison.Ordinal);
        Assert.Contains("Endnotes", fullText, StringComparison.Ordinal);
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_ProducesValidPackageAndStructuredDiagnostics()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            var result = DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            Assert.True(DocToDocxConverter.ValidatePackage(outputPath, out var validationError), validationError);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Message.Contains("Failed to read SEPX", StringComparison.Ordinal));

            using var archive = ZipFile.OpenRead(outputPath);
            using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
            var documentXml = reader.ReadToEnd();

            Assert.Contains("Text Formatting", documentXml, StringComparison.Ordinal);
            Assert.Contains("Inline formatting", documentXml, StringComparison.Ordinal);
            Assert.Contains("Footnotes", documentXml, StringComparison.Ordinal);
            Assert.Contains("Endnotes", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:pStyle w:val=\"Heading1\"", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:pStyle w:val=\"Heading2\"", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:vertAlign w:val=\"superscript\"", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:vertAlign w:val=\"subscript\"", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:u w:val=\"single\"", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:strike", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:tbl", documentXml, StringComparison.Ordinal);
            Assert.Contains("w:drawing", documentXml, StringComparison.Ordinal);
            Assert.DoesNotContain("\u0001", documentXml, StringComparison.Ordinal);
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_PreservesInlineFormattingFirstLineCharsIndent()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            using var archive = ZipFile.OpenRead(outputPath);
            var document = XDocument.Load(archive.GetEntry("word/document.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var paragraph = document
                .Descendants(w + "p")
                .FirstOrDefault(p => string.Concat(p.Descendants(w + "t").Select(t => (string?)t))
                    .Contains("Here, we demonstrate various types of inline text formatting", StringComparison.Ordinal));

            Assert.NotNull(paragraph);

            var indent = paragraph!
                .Element(w + "pPr")?
                .Element(w + "ind");

            Assert.NotNull(indent);
            Assert.Equal("206", (string?)indent!.Attribute(w + "firstLineChars"));
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void LoadDocument_Sample1Doc_PreservesStyledTextCharacterFormatting()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");

        var document = DocToDocxConverter.LoadDocument(inputPath);
        var paragraph = document.Paragraphs.First(p => p.Text.Contains("A paragraph with styled text:", StringComparison.Ordinal));
        var subtle = Assert.Single(paragraph.Runs.Where(run => run.Text.Contains("subtle emphasis", StringComparison.Ordinal)));
        var strong = Assert.Single(paragraph.Runs.Where(run => run.Text.Contains("strong text", StringComparison.Ordinal)));
        var intense = Assert.Single(paragraph.Runs.Where(run => run.Text.Contains("intense emphasis", StringComparison.Ordinal)));

        Assert.NotNull(subtle.Properties);
        Assert.True(subtle.Properties!.IsItalic);
        Assert.True(subtle.Properties.HasRgbColor);
        Assert.Equal(0xBD814Fu, subtle.Properties.RgbColor);
        Assert.Equal(24, subtle.Properties.FontSize);
        Assert.Equal(24, subtle.Properties.FontSizeCs);

        Assert.NotNull(strong.Properties);
        Assert.True(strong.Properties!.IsBold);
        Assert.Equal(24, strong.Properties.FontSize);
        Assert.Equal(24, strong.Properties.FontSizeCs);

        Assert.NotNull(intense.Properties);
        Assert.True(intense.Properties!.IsBold);
        Assert.True(intense.Properties.IsItalic);
        Assert.True(intense.Properties.HasRgbColor);
        Assert.Equal(0xBD814Fu, intense.Properties.RgbColor);
        Assert.Equal(24, intense.Properties.FontSize);
        Assert.Equal(24, intense.Properties.FontSizeCs);
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_PreservesStyledTextCharacterFormatting()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            using var archive = ZipFile.OpenRead(outputPath);
            var document = XDocument.Load(archive.GetEntry("word/document.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var paragraph = document
                .Descendants(w + "p")
                .FirstOrDefault(p => string.Concat(p.Descendants(w + "t").Select(t => (string?)t))
                    .Contains("A paragraph with styled text:", StringComparison.Ordinal));

            Assert.NotNull(paragraph);

            var subtle = FindRunByText(paragraph!, w, "subtle emphasis");
            var strong = FindRunByText(paragraph!, w, "strong text");
            var intense = FindRunByText(paragraph!, w, "intense emphasis");

            AssertRunFormatting(subtle, w, expectBold: false, expectItalic: true, expectColor: true);
            AssertRunFormatting(strong, w, expectBold: true, expectItalic: false, expectColor: false);
            AssertRunFormatting(intense, w, expectBold: true, expectItalic: true, expectColor: true);
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void LoadDocument_Sample1Doc_PreservesPrimaryHeadingFormatting()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");

        var document = DocToDocxConverter.LoadDocument(inputPath);
        var paragraph = document.Paragraphs.First(p => string.Equals(p.Text, "Text Formatting", StringComparison.Ordinal));

        Assert.NotNull(paragraph.Properties);
        Assert.Equal(ParagraphAlignment.Center, paragraph.Properties!.Alignment);

        var run = Assert.Single(paragraph.Runs.Where(r => r.Text.Contains("Text Formatting", StringComparison.Ordinal)));
        Assert.NotNull(run.Properties);
        Assert.True(run.Properties!.IsBold);
        Assert.True(run.Properties.HasRgbColor);
        Assert.Equal(0xBD814Fu, run.Properties.RgbColor);
        Assert.Equal(32, run.Properties.FontSize);
        Assert.Equal(32, run.Properties.FontSizeCs);
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_PreservesPrimaryHeadingFormatting()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            using var archive = ZipFile.OpenRead(outputPath);
            var document = XDocument.Load(archive.GetEntry("word/document.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var paragraph = document
                .Descendants(w + "p")
                .FirstOrDefault(p => string.Equals(string.Concat(p.Descendants(w + "t").Select(t => (string?)t)), "Text Formatting", StringComparison.Ordinal));

            Assert.NotNull(paragraph);
            Assert.Equal("center", (string?)paragraph!.Element(w + "pPr")?.Element(w + "jc")?.Attribute(w + "val"));

            var run = FindRunByText(paragraph!, w, "Text Formatting");
            var runProperties = run.Element(w + "rPr");
            Assert.NotNull(runProperties);
            Assert.NotNull(runProperties!.Element(w + "b"));
            Assert.Equal("4F81BD", (string?)runProperties.Element(w + "color")?.Attribute(w + "val"));
            Assert.Equal("32", (string?)runProperties.Element(w + "sz")?.Attribute(w + "val"));
            Assert.Equal("32", (string?)runProperties.Element(w + "szCs")?.Attribute(w + "val"));
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void LoadDocument_Sample1Doc_PreservesDocumentTitleStyle()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");

        var document = DocToDocxConverter.LoadDocument(inputPath);
        var paragraph = document.Paragraphs.First(p => string.Equals(p.Text, "Demonstration of DOCX support in calibre", StringComparison.Ordinal));

        Assert.NotNull(paragraph.Properties);
        var style = document.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == paragraph.Properties!.StyleIndex);
        Assert.NotNull(style);
        Assert.Equal("Title", style!.Name);
        Assert.Equal(ParagraphAlignment.Center, paragraph.Properties.Alignment);
        Assert.Equal("Ubuntu", style.RunProperties?.FontName);
        Assert.Contains(paragraph.Runs, run => string.Equals(run.Properties?.FontName, "Ubuntu", StringComparison.Ordinal));
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_PreservesDocumentTitleStyle()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            using var archive = ZipFile.OpenRead(outputPath);
            var document = XDocument.Load(archive.GetEntry("word/document.xml")!.Open());
            var styles = XDocument.Load(archive.GetEntry("word/styles.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var paragraph = document
                .Descendants(w + "p")
                .FirstOrDefault(p => string.Equals(string.Concat(p.Descendants(w + "t").Select(t => (string?)t)), "Demonstration of DOCX support in calibre", StringComparison.Ordinal));

            var titleStyle = styles
                .Descendants(w + "style")
                .FirstOrDefault(style => string.Equals((string?)style.Attribute(w + "styleId"), "Title", StringComparison.Ordinal));

            Assert.NotNull(paragraph);
            Assert.Equal("Title", (string?)paragraph!.Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val"));
            Assert.Equal("center", (string?)paragraph.Element(w + "pPr")?.Element(w + "jc")?.Attribute(w + "val"));
            Assert.Equal("Ubuntu", (string?)paragraph.Descendants(w + "rFonts").FirstOrDefault()?.Attribute(w + "ascii"));
            Assert.NotNull(titleStyle);
            Assert.Equal("Ubuntu", (string?)titleStyle!.Descendants(w + "rFonts").FirstOrDefault()?.Attribute(w + "ascii"));
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void LoadDocument_Sample1Doc_PreservesSimpleTableHeaderAndBodyRunFormatting()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");

        var document = DocToDocxConverter.LoadDocument(inputPath);
        var table = document.Tables.First(table => table.Rows.Count >= 2 && table.ColumnCount == 2 && string.Equals(table.Rows[0].Cells[0].Paragraphs[0].Text, "ITEM", StringComparison.Ordinal));

        var headerRun = Assert.Single(table.Rows[0].Cells[0].Paragraphs[0].Runs.Where(run => string.Equals(run.Text, "ITEM", StringComparison.Ordinal)));
        Assert.NotNull(headerRun.Properties);
        Assert.True(headerRun.Properties!.IsBold);
        Assert.True(headerRun.Properties.HasRgbColor);
        Assert.Equal(0xFFFFFFu, headerRun.Properties.RgbColor);

        var bodyRun = Assert.Single(table.Rows[1].Cells[0].Paragraphs[0].Runs.Where(run => string.Equals(run.Text, "Books", StringComparison.Ordinal)));
        Assert.NotNull(bodyRun.Properties);
        Assert.False(bodyRun.Properties!.IsBold);
        Assert.False(bodyRun.Properties.HasRgbColor && bodyRun.Properties.RgbColor == 0xFFFFFFu);
    }

    [Fact]
    public void LoadDocument_Sample1Doc_PreservesSimpleTableHeaderBackground()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");

        var document = DocToDocxConverter.LoadDocument(inputPath);
        var table = document.Tables.First(table => table.Rows.Count >= 2 && table.ColumnCount == 2 && string.Equals(table.Rows[0].Cells[0].Paragraphs[0].Text, "ITEM", StringComparison.Ordinal));

        Assert.All(table.Rows[0].Cells, cell =>
        {
            Assert.NotNull(cell.Properties);
            Assert.NotNull(cell.Properties!.Shading);
            Assert.NotEqual(0, cell.Properties.Shading!.BackgroundColor);
        });

        Assert.True(table.Rows[1].Cells.All(cell => cell.Properties?.Shading == null), "Body row should not inherit the header shading fallback.");
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_PreservesSimpleTableHeaderAndBodyRunFormatting()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            using var archive = ZipFile.OpenRead(outputPath);
            var document = XDocument.Load(archive.GetEntry("word/document.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var table = document
                .Descendants(w + "tbl")
                .First(tbl => string.Equals(string.Concat(tbl.Descendants(w + "t").Take(2).Select(t => (string?)t)), "ITEMNEEDED", StringComparison.Ordinal));

            var rows = table.Elements(w + "tr").ToList();
            Assert.True(rows.Count >= 2);

            var headerRunProperties = rows[0]
                .Elements(w + "tc")
                .First()
                .Descendants(w + "r")
                .First()
                .Element(w + "rPr");

            Assert.NotNull(headerRunProperties);
            Assert.NotNull(headerRunProperties!.Element(w + "b"));
            Assert.Equal("FFFFFF", (string?)headerRunProperties.Element(w + "color")?.Attribute(w + "val"));

            var bodyRunProperties = rows[1]
                .Elements(w + "tc")
                .First()
                .Descendants(w + "r")
                .First()
                .Element(w + "rPr");

            Assert.NotNull(bodyRunProperties);
            Assert.Null(bodyRunProperties!.Element(w + "b"));
            Assert.NotEqual("FFFFFF", (string?)bodyRunProperties.Element(w + "color")?.Attribute(w + "val"));
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_WritesSimpleTableHeaderShading()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            using var archive = ZipFile.OpenRead(outputPath);
            var document = XDocument.Load(archive.GetEntry("word/document.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var table = document
                .Descendants(w + "tbl")
                .First(tbl => string.Equals(string.Concat(tbl.Descendants(w + "t").Take(2).Select(t => (string?)t)), "ITEMNEEDED", StringComparison.Ordinal));

            var headerCells = table.Elements(w + "tr").First().Elements(w + "tc").ToList();
            Assert.NotEmpty(headerCells);

            Assert.All(headerCells, cell =>
            {
                var shading = cell.Element(w + "tcPr")?.Element(w + "shd");
                Assert.NotNull(shading);
                Assert.NotEqual("FFFFFF", (string?)shading!.Attribute(w + "fill"));
            });

            var firstBodyCellShading = table
                .Elements(w + "tr")
                .Skip(1)
                .First()
                .Elements(w + "tc")
                .First()
                .Element(w + "tcPr")?
                .Element(w + "shd");

            Assert.Null(firstBodyCellShading);
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void LoadDocument_Sample1Doc_PreservesDecember2007CalendarTableStructure()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");

        var document = DocToDocxConverter.LoadDocument(inputPath);
        var table = document.Tables.First(table =>
            table.Rows.Count > 0 &&
            table.Rows[0].Cells.Count > 0 &&
            string.Equals(table.Rows[0].Cells[0].Paragraphs[0].Text, "December 2007", StringComparison.Ordinal));

        Assert.Equal(13, table.ColumnCount);
        Assert.Equal(13, table.RowCount);

        var titleCell = Assert.Single(table.Rows[0].Cells);
        Assert.Equal(13, titleCell.ColumnSpan);
        Assert.Equal("December 2007", titleCell.Paragraphs[0].Text);

        Assert.Equal(13, table.Rows[1].Cells.Count);
        Assert.Equal(new[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" }, table.Rows[1].Cells.Where((_, index) => index % 2 == 0).Select(cell => cell.Paragraphs[0].Text));
        Assert.All(table.Rows[1].Cells.Where((_, index) => index % 2 == 1), cell => Assert.True(string.IsNullOrEmpty(cell.Paragraphs[0].Text)));

        Assert.Equal("1", table.Rows[2].Cells[12].Paragraphs[0].Text);
        Assert.All(table.Rows[3].Cells, cell => Assert.True(string.IsNullOrEmpty(cell.Paragraphs[0].Text)));
        Assert.Equal("30", table.Rows[12].Cells[0].Paragraphs[0].Text);
        Assert.Equal("31", table.Rows[12].Cells[2].Paragraphs[0].Text);
    }

    [Fact]
    public void ConvertWithWarnings_Sample1Doc_WritesDecember2007CalendarWithMergedTitleRow()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "sample1.doc");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

        try
        {
            DocToDocxConverter.ConvertWithWarnings(inputPath, outputPath);

            using var archive = ZipFile.OpenRead(outputPath);
            var document = XDocument.Load(archive.GetEntry("word/document.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var table = document
                .Descendants(w + "tbl")
                .First(tbl => string.Equals(string.Concat(tbl.Descendants(w + "t").Take(1).Select(t => (string?)t)), "December 2007", StringComparison.Ordinal));

            Assert.Equal(13, table.Element(w + "tblGrid")!.Elements(w + "gridCol").Count());

            var rows = table.Elements(w + "tr").ToList();
            Assert.Equal(13, rows.Count);

            var firstRowCells = rows[0].Elements(w + "tc").ToList();
            var mergedTitleCell = Assert.Single(firstRowCells);
            Assert.Equal("13", (string?)mergedTitleCell.Element(w + "tcPr")?.Element(w + "gridSpan")?.Attribute(w + "val"));
            Assert.Equal("December 2007", string.Concat(mergedTitleCell.Descendants(w + "t").Select(t => (string?)t)));

            Assert.Equal(13, rows[1].Elements(w + "tc").Count());

            var headerTexts = rows[1]
                .Elements(w + "tc")
                .Where((_, index) => index % 2 == 0)
                .Select(tc => string.Concat(tc.Descendants(w + "t").Select(t => (string?)t)))
                .ToList();
            Assert.Equal(new[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" }, headerTexts);

            var firstWeekTexts = rows[2]
                .Elements(w + "tc")
                .Select(tc => string.Concat(tc.Descendants(w + "t").Select(t => (string?)t)))
                .ToList();
            Assert.Equal("1", firstWeekTexts[12]);

            Assert.All(rows[3].Elements(w + "tc"), cell => Assert.Empty(string.Concat(cell.Descendants(w + "t").Select(t => (string?)t))));

            var lastWeekTexts = rows[12]
                .Elements(w + "tc")
                .Select(tc => string.Concat(tc.Descendants(w + "t").Select(t => (string?)t)))
                .ToList();
            Assert.Equal("30", lastWeekTexts[0]);
            Assert.Equal("31", lastWeekTexts[2]);

            var structuralHeadingInTable = table
                .Descendants(w + "t")
                .Any(text => string.Equals((string?)text, "Structural Elements", StringComparison.Ordinal));
            Assert.False(structuralHeadingInTable);
        }
        finally
        {
            DeleteIfExists(outputPath);
        }
    }

    [Fact]
    public void Convert_ProgressEvent_FiresAtLeastOnce()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");
        var updates = new List<ConversionProgress>();
        using var input = File.OpenRead(inputPath);
        using var outputStream = new MemoryStream();

        var progress = new ImmediateProgress<ConversionProgress>(updates.Add);
        DocToDocxConverter.Convert(input, outputStream, progress, password: null, enableHyperlinks: true, CancellationToken.None);
        Assert.NotEmpty(updates);
    }

    [Fact]
    public void ValidatePackage_StreamOverload_Works()
    {
        // create a temporary package on disk and then validate via a stream
        var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var document = new DocumentModel();
            document.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "x" } } });
            DocToDocxConverter.SaveDocument(document, tempPath);

            using var ms = new MemoryStream();
            using (var file = File.OpenRead(tempPath))
            {
                file.CopyTo(ms);
            }
            ms.Position = 0;
            Assert.True(DocToDocxConverter.ValidatePackage(ms, out var err));
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    [Fact]
    public void ColorHelper_ThemeResolution_Works()
    {
        var theme = new ThemeModel();
        theme.ColorMap["accent1"] = "112233";
        int colorValue = 0x01000000 | 4;
        Assert.Equal("112233", ColorHelper.ResolveThemeColorHex(colorValue, theme));
    }

    [Fact]
    public void SanitizeXmlString_ControlCharactersRemoved()
    {
        var method = typeof(Nedev.FileConverters.DocToDocx.Writers.DocumentWriter)
            .GetMethod("SanitizeXmlString", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        string input = "a\u0001b\u0007c";
        string cleaned = (string)method.Invoke(null, new object[] { input })!;
        Assert.Equal("abc", cleaned);
    }

    [Fact]
    public async Task Cli_HelpAndVersionBehave()
    {
        var writer = new StringWriter();
        var original = Console.Out;
        try
        {
            Console.SetOut(writer);
            await Nedev.FileConverters.DocToDocx.Cli.Program.Main(new[] { "-v" });
            var output = writer.ToString();
            Assert.Contains("Version", output);

            writer.GetStringBuilder().Clear();
            await Nedev.FileConverters.DocToDocx.Cli.Program.Main(new[] { "--help" });
            output = writer.ToString();
            Assert.Contains("--version", output);
        }
        finally
        {
            Console.SetOut(original);
        }
    }

    [Fact]
    public async Task Cli_UnknownOption_SetsExitCodeNonZero()
    {
        var writer = new StringWriter();
        var original = Console.Out;
        var originalExit = Environment.ExitCode;
        try
        {
            Console.SetOut(writer);
            await Nedev.FileConverters.DocToDocx.Cli.Program.Main(new[] { "input.doc", "output.docx", "--bogus" });
            Assert.Equal(1, Environment.ExitCode);
        }
        finally
        {
            Console.SetOut(original);
            Environment.ExitCode = originalExit;
        }
    }

    private static void WriteEntry(ZipArchive archive, string entryName, string content)
    {
        var entry = archive.CreateEntry(entryName);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

    private static XElement FindRunByText(XElement paragraph, XNamespace w, string text)
    {
        return paragraph
            .Elements(w + "r")
            .First(run => string.Concat(run.Descendants(w + "t").Select(t => (string?)t)).Contains(text, StringComparison.Ordinal));
    }

    private static void AssertRunFormatting(XElement run, XNamespace w, bool expectBold, bool expectItalic, bool expectColor)
    {
        var runProperties = run.Element(w + "rPr");
        Assert.NotNull(runProperties);

        Assert.Equal(expectBold, runProperties!.Element(w + "b") != null);
        Assert.Equal(expectItalic, runProperties.Element(w + "i") != null);
        Assert.Equal(expectColor, runProperties.Element(w + "color") != null);
        Assert.Null(runProperties.Element(w + "szCs"));
    }

    private sealed class ImmediateProgress<T> : IProgress<T>
    {
        private readonly Action<T> _handler;

        public ImmediateProgress(Action<T> handler)
        {
            _handler = handler;
        }

        public void Report(T value)
        {
            _handler(value);
        }
    }

    private sealed class NonSeekableReadOnlyStream : Stream
    {
        private readonly byte[] _buffer;
        private int _position;

        public NonSeekableReadOnlyStream(byte[] buffer)
        {
            _buffer = buffer;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var remaining = _buffer.Length - _position;
            if (remaining <= 0)
                return 0;

            var bytesToCopy = Math.Min(count, remaining);
            Buffer.BlockCopy(_buffer, _position, buffer, offset, bytesToCopy);
            _position += bytesToCopy;
            return bytesToCopy;
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}