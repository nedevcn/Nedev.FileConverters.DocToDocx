#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
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