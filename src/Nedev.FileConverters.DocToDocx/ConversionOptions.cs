using System;

namespace Nedev.FileConverters.DocToDocx;

/// <summary>
/// Options for controlling the DOC to DOCX conversion process.
/// </summary>
public sealed class ConversionOptions
{
    /// <summary>
    /// Gets or sets the password for encrypted documents.
    /// </summary>
    public string? Password { get; set; }

    /// <summary>
    /// Gets or sets whether to enable hyperlinks in the output document.
    /// Default is true.
    /// </summary>
    public bool EnableHyperlinks { get; set; } = true;

    /// <summary>
    /// Gets or sets whether to preserve embedded VBA macros.
    /// Default is false.
    /// </summary>
    public bool PreserveMacros { get; set; }

    /// <summary>
    /// Gets or sets whether to extract and embed external files.
    /// Default is true.
    /// </summary>
    public bool ExtractEmbeddedFiles { get; set; } = true;

    /// <summary>
    /// Gets or sets the maximum image resolution in DPI.
    /// Images with higher resolution will be downscaled.
    /// Set to 0 to disable limit.
    /// Default is 300.
    /// </summary>
    public int MaxImageResolution { get; set; } = 300;

    /// <summary>
    /// Gets or sets whether to skip package validation after conversion.
    /// Default is false.
    /// </summary>
    public bool SkipValidation { get; set; }

    /// <summary>
    /// Gets or sets whether to preserve revision tracking information.
    /// Default is true.
    /// </summary>
    public bool PreserveRevisions { get; set; } = true;

    /// <summary>
    /// Gets or sets whether to preserve comments and annotations.
    /// Default is true.
    /// </summary>
    public bool PreserveComments { get; set; } = true;

    /// <summary>
    /// Gets or sets the maximum nesting depth for tables.
    /// Default is 20.
    /// </summary>
    public int MaxTableNestingDepth { get; set; } = 20;

    /// <summary>
    /// Gets or sets the maximum file size in MB that can be processed.
    /// Files larger than this will throw an exception.
    /// Set to 0 to disable limit.
    /// Default is 100.
    /// </summary>
    public int MaxFileSizeMB { get; set; } = 100;

    /// <summary>
    /// Gets or sets whether to use strict OOXML conformance.
    /// When true, the output will strictly follow the OOXML specification.
    /// When false, some Word-specific extensions may be used for better compatibility.
    /// Default is false.
    /// </summary>
    public bool StrictConformance { get; set; }

    /// <summary>
    /// Gets or sets the timeout for the conversion operation.
    /// Default is 5 minutes.
    /// </summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Gets the default conversion options.
    /// </summary>
    public static ConversionOptions Default { get; } = new();

    /// <summary>
    /// Creates a copy of these options.
    /// </summary>
    public ConversionOptions Clone()
    {
        return new ConversionOptions
        {
            Password = Password,
            EnableHyperlinks = EnableHyperlinks,
            PreserveMacros = PreserveMacros,
            ExtractEmbeddedFiles = ExtractEmbeddedFiles,
            MaxImageResolution = MaxImageResolution,
            SkipValidation = SkipValidation,
            PreserveRevisions = PreserveRevisions,
            PreserveComments = PreserveComments,
            MaxTableNestingDepth = MaxTableNestingDepth,
            MaxFileSizeMB = MaxFileSizeMB,
            StrictConformance = StrictConformance,
            Timeout = Timeout
        };
    }

    /// <summary>
    /// Validates the options and throws an exception if any values are invalid.
    /// </summary>
    public void Validate()
    {
        if (MaxImageResolution < 0)
            throw new ArgumentOutOfRangeException(nameof(MaxImageResolution), MaxImageResolution, "MaxImageResolution must be non-negative.");

        if (MaxTableNestingDepth < 1)
            throw new ArgumentOutOfRangeException(nameof(MaxTableNestingDepth), MaxTableNestingDepth, "MaxTableNestingDepth must be at least 1.");

        if (MaxFileSizeMB < 0)
            throw new ArgumentOutOfRangeException(nameof(MaxFileSizeMB), MaxFileSizeMB, "MaxFileSizeMB must be non-negative.");

        if (Timeout <= TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(Timeout), Timeout, "Timeout must be positive.");
    }

    /// <summary>
    /// Validates that the file size is within the allowed limit.
    /// </summary>
    /// <param name="fileSize">The file size in bytes.</param>
    /// <exception cref="FileSizeLimitException">Thrown when file exceeds the limit.</exception>
    public void ValidateFileSize(long fileSize)
    {
        if (MaxFileSizeMB <= 0)
            return;

        var maxSizeBytes = (long)MaxFileSizeMB * 1024 * 1024;
        if (fileSize > maxSizeBytes)
        {
            throw new FileSizeLimitException(
                $"File size ({fileSize / 1024 / 1024}MB) exceeds the maximum allowed size ({MaxFileSizeMB}MB).",
                fileSize,
                maxSizeBytes);
        }
    }
}
