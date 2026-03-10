using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.Writers;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx;

/// <summary>
/// Main entry point for converting DOC files to DOCX
/// </summary>
public static class DocToDocxConverter
{
    private static readonly byte[] CompoundFileSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
    private static readonly byte[] ZipLocalHeaderSignature = { 0x50, 0x4B, 0x03, 0x04 };
    private static readonly string[] RequiredPackageEntries =
    {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml"
    };

    /// <summary>
    /// Converts a DOC file to DOCX format
    /// </summary>
    /// <param name="inputPath">Path to the input .doc file</param>
    /// <param name="outputPath">Path to the output .docx file</param>
    /// <param name="password">Optional password for encrypted DOC files.</param>
    /// <param name="enableHyperlinks">Whether hyperlink relationships should be emitted in the output DOCX.</param>
    public static void Convert(string inputPath, string outputPath, string? password = null, bool enableHyperlinks = true)
        => Convert(inputPath, outputPath, progress: null, password, enableHyperlinks, CancellationToken.None);

    /// <summary>
    /// Converts a DOC file to DOCX format and returns any non-fatal warnings captured during conversion.
    /// </summary>
    public static ConversionResult ConvertWithWarnings(string inputPath, string outputPath, string? password = null, bool enableHyperlinks = true)
        => ConvertWithWarnings(inputPath, outputPath, progress: null, password, enableHyperlinks, CancellationToken.None);

    /// <summary>
    /// Converts a DOC file to DOCX format asynchronously.
    /// </summary>
    public static Task ConvertAsync(string inputPath, string outputPath, string? password = null, bool enableHyperlinks = true, CancellationToken cancellationToken = default)
        => ConvertAsync(inputPath, outputPath, progress: null, password, enableHyperlinks, cancellationToken);

    /// <summary>
    /// Converts a DOC file to DOCX format asynchronously and returns captured non-fatal warnings.
    /// </summary>
    public static Task<ConversionResult> ConvertWithWarningsAsync(string inputPath, string outputPath, string? password = null, bool enableHyperlinks = true, CancellationToken cancellationToken = default)
        => ConvertWithWarningsAsync(inputPath, outputPath, progress: null, password, enableHyperlinks, cancellationToken);

    /// <summary>
    /// Converts a DOC file to DOCX format asynchronously with progress reporting.
    /// </summary>
    public static Task ConvertAsync(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password = null, bool enableHyperlinks = true, CancellationToken cancellationToken = default)
        => Task.Run(() => Convert(inputPath, outputPath, progress, password, enableHyperlinks, cancellationToken), cancellationToken);

    /// <summary>
    /// Converts a DOC file to DOCX format asynchronously with progress reporting and captured warnings.
    /// </summary>
    public static Task<ConversionResult> ConvertWithWarningsAsync(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password = null, bool enableHyperlinks = true, CancellationToken cancellationToken = default)
        => Task.Run(() => ConvertWithWarnings(inputPath, outputPath, progress, password, enableHyperlinks, cancellationToken), cancellationToken);

    /// <summary>
    /// Converts a DOC file to DOCX format with progress reporting
    /// </summary>
    public static void Convert(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password = null, bool enableHyperlinks = true)
        => Convert(inputPath, outputPath, progress, password, enableHyperlinks, CancellationToken.None);

    /// <summary>
    /// Converts a DOC file to DOCX format with progress reporting and returns captured warnings.
    /// </summary>
    public static ConversionResult ConvertWithWarnings(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password = null, bool enableHyperlinks = true)
        => ConvertWithWarnings(inputPath, outputPath, progress, password, enableHyperlinks, CancellationToken.None);

    /// <summary>
    /// Converts a DOC file to DOCX format with progress reporting and cancellation.
    /// </summary>
    public static void Convert(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password, bool enableHyperlinks, CancellationToken cancellationToken)
    {
        ValidateConversionPaths(inputPath, outputPath);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            Report(progress, ConversionStage.Initializing, 0, $"Inspecting input '{inputPath}'.");

            var inputKind = DetectInputKind(inputPath);
            if (inputKind == InputDocumentKind.Docx)
            {
                CopyDocxInput(inputPath, outputPath);
                Report(progress, ConversionStage.Complete, 100, $"Copied DOCX package to '{outputPath}'.");
                return;
            }

            if (inputKind != InputDocumentKind.Doc)
            {
                throw new InvalidDataException($"Unsupported input format for '{inputPath}'. Expected a DOC compound file or a DOCX ZIP package.");
            }

            Report(progress, ConversionStage.Reading, 15, $"Reading DOC file '{inputPath}'.");
            using var reader = new DocReader(inputPath, password);
            reader.Load();
            cancellationToken.ThrowIfCancellationRequested();

            var document = reader.Document;
            var imageBytes = document.Images.Sum(image => image.Data?.Length ?? 0);
            Report(progress, ConversionStage.Parsing, 55,
                $"Recovered {document.Paragraphs.Count} paragraphs, {document.Tables.Count} tables, {document.Images.Count} images ({imageBytes / 1024} KB).");

            WriteDocumentPackage(document, outputPath, enableHyperlinks, validatePackage: true, cancellationToken);
            Report(progress, ConversionStage.Complete, 100, $"Wrote DOCX package to '{outputPath}'.");
        }
        catch (OperationCanceledException)
        {
            Report(progress, ConversionStage.Error, 100, "Conversion canceled.");
            throw;
        }
        catch (Exception ex)
        {
            Report(progress, ConversionStage.Error, 100, ex.Message);
            throw;
        }
    }

    /// <summary>
    /// Converts a DOC file to DOCX format with progress reporting, cancellation, and captured warnings.
    /// </summary>
    public static ConversionResult ConvertWithWarnings(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password, bool enableHyperlinks, CancellationToken cancellationToken)
    {
        var diagnostics = new List<ConversionDiagnostic>();
        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            Convert(inputPath, outputPath, progress, password, enableHyperlinks, cancellationToken);
        }

        return new ConversionResult(outputPath, diagnostics);
    }
    
    /// <summary>
    /// Loads a DOC file and returns the document model
    /// </summary>
    public static DocumentModel LoadDocument(string inputPath, string? password = null)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
            throw new ArgumentException("Input path cannot be null or empty.", nameof(inputPath));

        using var reader = new DocReader(inputPath, password);
        reader.Load();
        return reader.Document;
    }
    
    /// <summary>
    /// Saves a document model to DOCX format
    /// </summary>
    public static void SaveDocument(DocumentModel document, string outputPath) =>
        SaveDocument(document, outputPath, enableHyperlinks: true);

    /// <summary>
    /// Saves a document model to DOCX format, optionally disabling hyperlink
    /// relationships.  Turning off hyperlinks prevents Word from warning about
    /// fields linking to other files when the document is opened.
    /// </summary>
    public static void SaveDocument(DocumentModel document, string outputPath, bool enableHyperlinks)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));

        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path cannot be null or empty.", nameof(outputPath));

        WriteDocumentPackage(document, outputPath, enableHyperlinks, validatePackage: true, CancellationToken.None);
    }

    /// <summary>
    /// Performs simple validation of a generated DOCX package: each XML part must parse.
    /// Returns true if all XML entries are well-formed; otherwise false and an error message.
    /// </summary>
    public static bool ValidatePackage(string path, out string? errorMessage)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            errorMessage = "Package path cannot be null or empty.";
            return false;
        }

        if (!File.Exists(path))
        {
            errorMessage = $"Package file '{path}' does not exist.";
            return false;
        }

        try
        {
            using var stream = File.OpenRead(path);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var entryNames = archive.Entries
                .Select(entry => NormalizeEntryPath(entry.FullName))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var requiredEntry in RequiredPackageEntries)
            {
                if (!entryNames.Contains(requiredEntry))
                {
                    errorMessage = $"Missing required package part '{requiredEntry}'.";
                    return false;
                }
            }

            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = XmlReader.Create(entry.Open(), CreateXmlReaderSettings());
                    while (reader.Read()) { }
                }

                if (!entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    continue;

                var missingTarget = FindMissingInternalRelationshipTarget(archive, entry);
                if (missingTarget != null)
                {
                    errorMessage = $"Relationship part '{entry.FullName}' references missing target '{missingTarget}'.";
                    return false;
                }
            }

            errorMessage = null;
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            return false;
        }
    }

    private static void ValidateConversionPaths(string inputPath, string outputPath)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
            throw new ArgumentException("Input path cannot be null or empty.", nameof(inputPath));

        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path cannot be null or empty.", nameof(outputPath));

        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Input file '{inputPath}' was not found.", inputPath);

        var inputFullPath = Path.GetFullPath(inputPath);
        var outputFullPath = Path.GetFullPath(outputPath);
        if (string.Equals(inputFullPath, outputFullPath, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Input and output paths must be different.", nameof(outputPath));
    }

    private static void WriteDocumentPackage(DocumentModel document, string outputPath, bool enableHyperlinks, bool validatePackage, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureOutputDirectory(outputPath);

        using (var stream = File.Create(outputPath))
        {
            using var zipWriter = new ZipWriter(stream);
            var options = new Writers.DocumentWriterOptions { EnableHyperlinks = enableHyperlinks };
            Logger.Info($"DocToDocxConverter.WriteDocumentPackage START: output={outputPath} paragraphs={document.Paragraphs.Count} tables={document.Tables.Count}");
            zipWriter.WriteDocument(document, options);
            Logger.Info($"DocToDocxConverter.WriteDocumentPackage AFTER WriteDocument: output={outputPath}");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (validatePackage && !ValidatePackage(outputPath, out var validationMessage))
            throw new InvalidDataException($"Generated DOCX package failed validation: {validationMessage}");
    }

    private static void CopyDocxInput(string inputPath, string outputPath)
    {
        EnsureOutputDirectory(outputPath);
        File.Copy(inputPath, outputPath, overwrite: true);
    }

    private static void EnsureOutputDirectory(string outputPath)
    {
        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);
    }

    private static void Report(IProgress<ConversionProgress>? progress, ConversionStage stage, int percentComplete, string message)
    {
        progress?.Report(new ConversionProgress
        {
            Stage = stage,
            PercentComplete = percentComplete,
            Message = message
        });
    }

    private static InputDocumentKind DetectInputKind(string inputPath)
    {
        using var stream = File.OpenRead(inputPath);
        return DetectInputKind(stream, Path.GetExtension(inputPath));
    }

    private static InputDocumentKind DetectInputKind(Stream stream, string? extensionHint)
    {
        Span<byte> header = stackalloc byte[8];
        var originalPosition = stream.CanSeek ? stream.Position : 0;
        try
        {
            var bytesRead = stream.Read(header);
            if (bytesRead >= ZipLocalHeaderSignature.Length && header[..ZipLocalHeaderSignature.Length].SequenceEqual(ZipLocalHeaderSignature))
                return InputDocumentKind.Docx;

            if (bytesRead >= CompoundFileSignature.Length && header[..CompoundFileSignature.Length].SequenceEqual(CompoundFileSignature))
                return InputDocumentKind.Doc;
        }
        finally
        {
            if (stream.CanSeek)
                stream.Position = originalPosition;
        }

        return extensionHint?.ToLowerInvariant() switch
        {
            ".docx" => InputDocumentKind.Docx,
            ".doc" => InputDocumentKind.Doc,
            _ => InputDocumentKind.Unknown
        };
    }

    private static XmlReaderSettings CreateXmlReaderSettings()
    {
        return new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            IgnoreWhitespace = true,
            CloseInput = true
        };
    }

    private static string? FindMissingInternalRelationshipTarget(ZipArchive archive, ZipArchiveEntry relationshipsEntry)
    {
        using var stream = relationshipsEntry.Open();
        var relationships = XDocument.Load(stream, LoadOptions.None);
        if (relationships.Root == null)
            return relationshipsEntry.FullName;

        XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        var entryNames = archive.Entries
            .Select(entry => NormalizeEntryPath(entry.FullName))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var relationship in relationships.Root.Elements(relationshipsNamespace + "Relationship"))
        {
            var targetMode = (string?)relationship.Attribute("TargetMode");
            if (string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase))
                continue;

            var target = (string?)relationship.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
                return "<empty>";

            var resolvedTarget = ResolveRelationshipTarget(relationshipsEntry.FullName, target);
            if (!entryNames.Contains(resolvedTarget))
                return resolvedTarget;
        }

        return null;
    }

    private static string ResolveRelationshipTarget(string relationshipsPath, string target)
    {
        var normalizedTarget = NormalizeEntryPath(target);
        if (target.Length > 0 && target[0] == '/')
            return normalizedTarget;

        var sourceDirectory = GetSourcePartDirectory(relationshipsPath);
        var segments = new List<string>();
        if (!string.IsNullOrEmpty(sourceDirectory))
            segments.AddRange(sourceDirectory.Split('/', StringSplitOptions.RemoveEmptyEntries));

        foreach (var segment in normalizedTarget.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
                continue;

            if (segment == "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);

                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }

    private static string GetSourcePartDirectory(string relationshipsPath)
    {
        var normalizedPath = NormalizeEntryPath(relationshipsPath);
        if (string.Equals(normalizedPath, "_rels/.rels", StringComparison.OrdinalIgnoreCase))
            return string.Empty;

        var relationsDirectory = Path.GetDirectoryName(normalizedPath)?.Replace('\\', '/') ?? string.Empty;
        const string marker = "/_rels";
        var markerIndex = relationsDirectory.LastIndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex >= 0)
            return relationsDirectory.Substring(0, markerIndex);

        return relationsDirectory;
    }

    private static string NormalizeEntryPath(string path)
    {
        return path.Replace('\\', '/').TrimStart('/');
    }
}

internal enum InputDocumentKind
{
    Unknown,
    Doc,
    Docx
}

/// <summary>
/// Represents the progress of a document conversion
/// </summary>
public class ConversionProgress
{
    /// <summary>
    /// Current conversion stage.
    /// </summary>
    public ConversionStage Stage { get; set; }

    /// <summary>
    /// Percentage reported for the current conversion stage.
    /// </summary>
    public int PercentComplete { get; set; }

    /// <summary>
    /// Human-readable status message for the current progress update.
    /// </summary>
    public string? Message { get; set; }
}

/// <summary>
/// Represents the output path and captured non-fatal diagnostics for a conversion.
/// </summary>
public sealed class ConversionResult
{
    public ConversionResult(string outputPath, IReadOnlyList<ConversionDiagnostic> diagnostics)
    {
        OutputPath = outputPath ?? throw new ArgumentNullException(nameof(outputPath));
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        Warnings = diagnostics.Select(static diagnostic => diagnostic.FormattedMessage).ToArray();
    }

    public ConversionResult(string outputPath, IReadOnlyList<string> warnings)
    {
        OutputPath = outputPath ?? throw new ArgumentNullException(nameof(outputPath));
        Warnings = warnings ?? throw new ArgumentNullException(nameof(warnings));
        Diagnostics = warnings.Select(static warning =>
            new ConversionDiagnostic(DateTime.UtcNow, Logger.LogLevel.Warning, warning, warning, exceptionType: null, exceptionMessage: null)).ToArray();
    }

    public string OutputPath { get; }
    public IReadOnlyList<ConversionDiagnostic> Diagnostics { get; }
    public IReadOnlyList<string> Warnings { get; }
}

/// <summary>
/// Represents the current stage of document conversion
/// </summary>
public enum ConversionStage
{
    /// <summary>
    /// Inspecting inputs and preparing the conversion pipeline.
    /// </summary>
    Initializing,

    /// <summary>
    /// Reading the source document streams.
    /// </summary>
    Reading,

    /// <summary>
    /// Translating binary structures into the document model.
    /// </summary>
    Parsing,

    /// <summary>
    /// Writing the output DOCX package.
    /// </summary>
    Writing,

    /// <summary>
    /// Conversion completed successfully.
    /// </summary>
    Complete,

    /// <summary>
    /// Conversion stopped because of an error or cancellation.
    /// </summary>
    Error
}



/// <summary>
/// DOCX Writer - Main writer class that orchestrates the output
/// </summary>
public class DocxWriter : IDisposable
{
    private readonly Stream _outputStream;
    private readonly ZipWriter _zipWriter;
    private readonly bool _ownsStream;
    
    /// <summary>
    /// Initializes a writer that emits a DOCX package to an existing stream.
    /// </summary>
    /// <param name="outputStream">Destination stream for the generated DOCX package.</param>
    public DocxWriter(Stream outputStream)
    {
        _outputStream = outputStream;
        _zipWriter = new ZipWriter(outputStream);
        _ownsStream = false;
    }
    
    /// <summary>
    /// Initializes a writer that creates a DOCX package at the specified path.
    /// </summary>
    /// <param name="outputPath">Destination file path for the generated DOCX package.</param>
    public DocxWriter(string outputPath)
    {
        _outputStream = File.Create(outputPath);
        try
        {
            _zipWriter = new ZipWriter(_outputStream);
        }
        catch
        {
            _outputStream.Dispose();
            throw;
        }
        _ownsStream = true;
    }
    
    /// <summary>
    /// Writes the document to DOCX format
    /// </summary>
    public void Write(DocumentModel document)
    {
        _zipWriter.WriteDocument(document);
    }
    
    /// <summary>
    /// Disposes the writer
    /// </summary>
    public void Dispose()
    {
        _zipWriter?.Dispose();
        if (_ownsStream)
        {
            _outputStream?.Dispose();
        }
    }
}
