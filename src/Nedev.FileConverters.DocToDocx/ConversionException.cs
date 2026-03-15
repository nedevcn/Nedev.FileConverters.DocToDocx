using System;

namespace Nedev.FileConverters.DocToDocx;

/// <summary>
/// Base exception for all conversion-related errors.
/// </summary>
public class ConversionException : Exception
{
    /// <summary>
    /// Gets the type of conversion error.
    /// </summary>
    public ConversionErrorType ErrorType { get; }

    /// <summary>
    /// Gets the input file path, if available.
    /// </summary>
    public string? InputPath { get; }

    public ConversionException(ConversionErrorType errorType, string message)
        : base(message)
    {
        ErrorType = errorType;
    }

    public ConversionException(ConversionErrorType errorType, string message, Exception innerException)
        : base(message, innerException)
    {
        ErrorType = errorType;
    }

    public ConversionException(ConversionErrorType errorType, string message, string? inputPath)
        : base(message)
    {
        ErrorType = errorType;
        InputPath = inputPath;
    }

    public ConversionException(ConversionErrorType errorType, string message, string? inputPath, Exception innerException)
        : base(message, innerException)
    {
        ErrorType = errorType;
        InputPath = inputPath;
    }
}

/// <summary>
/// Exception thrown when the input file format is not supported.
/// </summary>
public sealed class UnsupportedFormatException : ConversionException
{
    public UnsupportedFormatException(string message)
        : base(ConversionErrorType.UnsupportedFormat, message)
    {
    }

    public UnsupportedFormatException(string message, Exception innerException)
        : base(ConversionErrorType.UnsupportedFormat, message, innerException)
    {
    }
}

/// <summary>
/// Exception thrown when the input file is corrupted or invalid.
/// </summary>
public sealed class CorruptedFileException : ConversionException
{
    public CorruptedFileException(string message)
        : base(ConversionErrorType.CorruptedFile, message)
    {
    }

    public CorruptedFileException(string message, Exception innerException)
        : base(ConversionErrorType.CorruptedFile, message, innerException)
    {
    }
}

/// <summary>
/// Exception thrown when a password is required or incorrect.
/// </summary>
public sealed class EncryptionException : ConversionException
{
    public EncryptionException(string message)
        : base(ConversionErrorType.EncryptionError, message)
    {
    }

    public EncryptionException(string message, Exception innerException)
        : base(ConversionErrorType.EncryptionError, message, innerException)
    {
    }
}

/// <summary>
/// Exception thrown when the file size exceeds the allowed limit.
/// </summary>
public sealed class FileSizeLimitException : ConversionException
{
    public long FileSize { get; }
    public long MaxSize { get; }

    public FileSizeLimitException(string message, long fileSize, long maxSize)
        : base(ConversionErrorType.FileTooLarge, message)
    {
        FileSize = fileSize;
        MaxSize = maxSize;
    }
}

/// <summary>
/// Exception thrown when the conversion times out.
/// </summary>
public sealed class ConversionTimeoutException : ConversionException
{
    public TimeSpan Timeout { get; }

    public ConversionTimeoutException(string message, TimeSpan timeout)
        : base(ConversionErrorType.Timeout, message)
    {
        Timeout = timeout;
    }
}

/// <summary>
/// Exception thrown when the output is invalid or validation fails.
/// </summary>
public sealed class ValidationException : ConversionException
{
    public string? ValidationError { get; }

    public ValidationException(string message, string? validationError = null)
        : base(ConversionErrorType.ValidationFailed, message)
    {
        ValidationError = validationError;
    }
}

/// <summary>
/// Types of conversion errors.
/// </summary>
public enum ConversionErrorType
{
    /// <summary>Unknown error type.</summary>
    Unknown,

    /// <summary>The file format is not supported.</summary>
    UnsupportedFormat,

    /// <summary>The file is corrupted or invalid.</summary>
    CorruptedFile,

    /// <summary>Encryption-related error (password required or incorrect).</summary>
    EncryptionError,

    /// <summary>The file is too large.</summary>
    FileTooLarge,

    /// <summary>The conversion timed out.</summary>
    Timeout,

    /// <summary>Output validation failed.</summary>
    ValidationFailed,

    /// <summary>Out of memory during conversion.</summary>
    OutOfMemory,

    /// <summary>Permission denied accessing file.</summary>
    PermissionDenied,

    /// <summary>Network or IO error.</summary>
    IOError
}
