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
    public static void Convert(string inputPath, string outputPath, string? password = null)
    {
        using var reader = new DocReader(inputPath, password);
        
        Console.WriteLine($"Reading document: {inputPath}");
        reader.Load();
        var doc = reader.Document;
        var imageBytes = doc.Images.Sum(i => i.Data?.Length ?? 0);
        Console.WriteLine($"Parsed {doc.Paragraphs.Count} paragraphs, {doc.Tables.Count} tables, {doc.Images.Count} images ({imageBytes / 1024} KB)");
        
        // Ensure output directory exists
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }
        
        Console.WriteLine($"Writing document: {outputPath}");
        
        using var stream = File.Create(outputPath);
        using var zipWriter = new ZipWriter(stream);
        zipWriter.WriteDocument(reader.Document);
        
        Console.WriteLine("Conversion complete!");
    }
    
    /// <summary>
    /// Converts a DOC file to DOCX format asynchronously
    /// </summary>
    public static async Task ConvertAsync(string inputPath, string outputPath, string? password = null, CancellationToken cancellationToken = default)
    {
        await Task.Run(() => Convert(inputPath, outputPath, password), cancellationToken);
    }
    
    /// <summary>
    /// Converts a DOC file to DOCX format with progress reporting
    /// </summary>
    public static void Convert(string inputPath, string outputPath, IProgress<ConversionProgress>? progress, string? password = null)
    {
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
        
        using var stream = File.Create(outputPath);
        using var zipWriter = new ZipWriter(stream);
        
        progress?.Report(new ConversionProgress { Stage = ConversionStage.Writing, PercentComplete = 80 });
        zipWriter.WriteDocument(reader.Document);
        
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
    public static void SaveDocument(DocumentModel document, string outputPath)
    {
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }
        
        using var stream = File.Create(outputPath);
        using var zipWriter = new ZipWriter(stream);
        zipWriter.WriteDocument(document);
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
