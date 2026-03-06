using System.IO;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Readers;
using Nedev.DocToDocx.Writers;

namespace Nedev.DocToDocx;

/// <summary>
/// Main entry point for converting DOC files to DOCX
/// </summary>
public static class DocToDocxConverter
{
    /// <summary>
    /// Converts a DOC file to DOCX format
    /// </summary>
    /// <param name="inputPath">Path to the input .doc file</param>
    /// <param name="outputPath">Path to the output .docx file</param>
    public static void Convert(string inputPath, string outputPath, string? password = null, bool enableHyperlinks = true)
    {
        // if the input is already a DOCX file just copy it (CLI supports this, API should too)
        if (Path.GetExtension(inputPath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            // ensure output directory exists
            var outDir1 = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir1) && !Directory.Exists(outDir1))
            {
                Directory.CreateDirectory(outDir1);
            }
            File.Copy(inputPath, outputPath, overwrite: true);
            return;
        }

        using var reader = new DocReader(inputPath, password);
        
        Console.WriteLine($"Reading document: {inputPath}");
        reader.Load();
        var doc = reader.Document;
        var imageBytes = doc.Images.Sum(i => i.Data?.Length ?? 0);
        Console.WriteLine($"Parsed {doc.Paragraphs.Count} paragraphs, {doc.Tables.Count} tables, {doc.Images.Count} images ({imageBytes / 1024} KB)");
        
        // Ensure output directory exists
        var outputDir2 = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir2) && !Directory.Exists(outputDir2))
        {
            Directory.CreateDirectory(outputDir2);
        }
        
        Console.WriteLine($"Writing document: {outputPath}");
        
        using var stream = File.Create(outputPath);
        using var zipWriter = new ZipWriter(stream);
        var options = new Writers.DocumentWriterOptions { EnableHyperlinks = enableHyperlinks };
        zipWriter.WriteDocument(reader.Document, options);
        
        // Explicitly dispose the writer to flush the ZIP central directory
        // before the underlying stream is closed or before we return.
        zipWriter.Dispose();
        
        // validate output to catch accidental corruption early
        if (!ValidatePackage(outputPath, out var validationMessage))
        {
            Console.WriteLine("Warning: generated DOCX failed validation: " + validationMessage);
        }
        else
        {
            Console.WriteLine("Conversion complete!");
        }
    }
    
    /// <summary>
    /// Converts a DOC file to DOCX format asynchronously
    /// </summary>
    public static async Task ConvertAsync(string inputPath, string outputPath, string? password = null, bool enableHyperlinks = true, CancellationToken cancellationToken = default)
    {
        await Task.Run(() => Convert(inputPath, outputPath, password, enableHyperlinks), cancellationToken);
    }
    
    /// <summary>
    /// Converts a DOC file to DOCX format with progress reporting
    /// </summary>
    public static void Convert(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password = null, bool enableHyperlinks = true)
    {
        // docx input should simply be copied, but still report stages for compatibility
        if (Path.GetExtension(inputPath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            progress?.Report(new ConversionProgress { Stage = ConversionStage.Reading, PercentComplete = 0 });
            progress?.Report(new ConversionProgress { Stage = ConversionStage.Reading, PercentComplete = 20 });
            // ensure output dir exists
            var outDir3 = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir3) && !Directory.Exists(outDir3))
            {
                Directory.CreateDirectory(outDir3);
            }
            progress?.Report(new ConversionProgress { Stage = ConversionStage.Writing, PercentComplete = 60 });
            File.Copy(inputPath, outputPath, overwrite: true);
            progress?.Report(new ConversionProgress { Stage = ConversionStage.Writing, PercentComplete = 80 });
            progress?.Report(new ConversionProgress { Stage = ConversionStage.Complete, PercentComplete = 100 });
            return;
        }

        progress?.Report(new ConversionProgress { Stage = ConversionStage.Reading, PercentComplete = 0 });
        
        using var reader = new DocReader(inputPath, password);
        
        progress?.Report(new ConversionProgress { Stage = ConversionStage.Reading, PercentComplete = 20 });
        reader.Load();
        
        progress?.Report(new ConversionProgress { Stage = ConversionStage.Reading, PercentComplete = 40 });
        
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }
        
        progress?.Report(new ConversionProgress { Stage = ConversionStage.Writing, PercentComplete = 60 });
        
        using (var stream = File.Create(outputPath))
        {
            using var zipWriter = new ZipWriter(stream);

            progress?.Report(new ConversionProgress { Stage = ConversionStage.Writing, PercentComplete = 80 });
            var options = new Writers.DocumentWriterOptions { EnableHyperlinks = enableHyperlinks };
            zipWriter.WriteDocument(reader.Document, options);

            // Explicitly dispose the writer to flush the ZIP central directory
            zipWriter.Dispose();
        }

        // validate after stream closed
        if (!ValidatePackage(outputPath, out var validationMessage))
        {
            Console.WriteLine("Warning: generated DOCX failed validation: " + validationMessage);
        }
        
        progress?.Report(new ConversionProgress { Stage = ConversionStage.Complete, PercentComplete = 100 });
    }
    
    /// <summary>
    /// Loads a DOC file and returns the document model
    /// </summary>
    public static DocumentModel LoadDocument(string inputPath, string? password = null)
    {
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
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }
        
        using (var stream = File.Create(outputPath))
        {
            using var zipWriter = new ZipWriter(stream);
            var options = new Writers.DocumentWriterOptions { EnableHyperlinks = enableHyperlinks };
            zipWriter.WriteDocument(document, options);
            // Explicitly dispose the writer to flush the ZIP central directory
            zipWriter.Dispose();
        }
    }

    /// <summary>
    /// Performs simple validation of a generated DOCX package: each XML part must parse.
    /// Returns true if all XML entries are well-formed; otherwise false and an error message.
    /// </summary>
    public static bool ValidatePackage(string path, out string? errorMessage)
    {
        try
        {
            using var archive = new System.IO.Compression.ZipArchive(File.OpenRead(path), System.IO.Compression.ZipArchiveMode.Read);
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = System.Xml.XmlReader.Create(entry.Open());
                    while (reader.Read()) { }
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
}
/// <summary>
/// Represents the progress of a document conversion
/// </summary>
public class ConversionProgress
{
    public ConversionStage Stage { get; set; }
    public int PercentComplete { get; set; }
    public string? Message { get; set; }
}

/// <summary>
/// Represents the current stage of document conversion
/// </summary>
public enum ConversionStage
{
    Initializing,
    Reading,
    Parsing,
    Writing,
    Complete,
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
    
    public DocxWriter(Stream outputStream)
    {
        _outputStream = outputStream;
        _zipWriter = new ZipWriter(outputStream);
        _ownsStream = false;
    }
    
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
