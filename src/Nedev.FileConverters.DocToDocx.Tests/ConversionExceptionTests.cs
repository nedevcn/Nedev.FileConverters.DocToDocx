using System;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class ConversionExceptionTests
{
    [Fact]
    public void ConversionException_WithErrorType_SetsProperties()
    {
        var exception = new ConversionException(ConversionErrorType.CorruptedFile, "File is corrupted");

        Assert.Equal(ConversionErrorType.CorruptedFile, exception.ErrorType);
        Assert.Equal("File is corrupted", exception.Message);
    }

    [Fact]
    public void ConversionException_WithInnerException_SetsProperties()
    {
        var inner = new InvalidOperationException("Inner error");
        var exception = new ConversionException(ConversionErrorType.IOError, "Conversion failed", inner);

        Assert.Equal(ConversionErrorType.IOError, exception.ErrorType);
        Assert.Same(inner, exception.InnerException);
    }

    [Fact]
    public void ConversionException_WithInputPath_SetsProperties()
    {
        var exception = new ConversionException(ConversionErrorType.UnsupportedFormat, "Not supported", "test.doc");

        Assert.Equal(ConversionErrorType.UnsupportedFormat, exception.ErrorType);
        Assert.Equal("test.doc", exception.InputPath);
    }

    [Fact]
    public void UnsupportedFormatException_SetsCorrectErrorType()
    {
        var exception = new UnsupportedFormatException("Format not supported");

        Assert.Equal(ConversionErrorType.UnsupportedFormat, exception.ErrorType);
    }

    [Fact]
    public void CorruptedFileException_SetsCorrectErrorType()
    {
        var exception = new CorruptedFileException("File corrupted");

        Assert.Equal(ConversionErrorType.CorruptedFile, exception.ErrorType);
    }

    [Fact]
    public void EncryptionException_SetsCorrectErrorType()
    {
        var exception = new EncryptionException("Wrong password");

        Assert.Equal(ConversionErrorType.EncryptionError, exception.ErrorType);
    }

    [Fact]
    public void FileSizeLimitException_SetsSizeProperties()
    {
        var exception = new FileSizeLimitException("Too large", 2000000, 1000000);

        Assert.Equal(ConversionErrorType.FileTooLarge, exception.ErrorType);
        Assert.Equal(2000000, exception.FileSize);
        Assert.Equal(1000000, exception.MaxSize);
    }

    [Fact]
    public void ConversionTimeoutException_SetsTimeout()
    {
        var timeout = TimeSpan.FromMinutes(2);
        var exception = new ConversionTimeoutException("Timed out", timeout);

        Assert.Equal(ConversionErrorType.Timeout, exception.ErrorType);
        Assert.Equal(timeout, exception.Timeout);
    }

    [Fact]
    public void ValidationException_SetsValidationError()
    {
        var exception = new ValidationException("Validation failed", "Missing required part");

        Assert.Equal(ConversionErrorType.ValidationFailed, exception.ErrorType);
        Assert.Equal("Missing required part", exception.ValidationError);
    }
}
