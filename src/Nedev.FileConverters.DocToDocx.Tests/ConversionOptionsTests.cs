using System;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class ConversionOptionsTests
{
    [Fact]
    public void DefaultOptions_AreValid()
    {
        var options = ConversionOptions.Default;
        options.Validate();

        Assert.True(options.EnableHyperlinks);
        Assert.False(options.PreserveMacros);
        Assert.True(options.ExtractEmbeddedFiles);
        Assert.Equal(300, options.MaxImageResolution);
        Assert.False(options.SkipValidation);
        Assert.True(options.PreserveRevisions);
        Assert.True(options.PreserveComments);
        Assert.Equal(20, options.MaxTableNestingDepth);
        Assert.Equal(100, options.MaxFileSizeMB);
        Assert.False(options.StrictConformance);
        Assert.Equal(TimeSpan.FromMinutes(5), options.Timeout);
    }

    [Fact]
    public void Clone_CreatesIndependentCopy()
    {
        var original = new ConversionOptions
        {
            Password = "secret",
            EnableHyperlinks = false,
            MaxImageResolution = 150
        };

        var clone = original.Clone();

        Assert.Equal(original.Password, clone.Password);
        Assert.Equal(original.EnableHyperlinks, clone.EnableHyperlinks);
        Assert.Equal(original.MaxImageResolution, clone.MaxImageResolution);

        // Modify clone should not affect original
        clone.Password = "newpassword";
        Assert.Equal("secret", original.Password);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-100)]
    public void Validate_InvalidMaxImageResolution_Throws(int value)
    {
        var options = new ConversionOptions { MaxImageResolution = value };
        Assert.Throws<ArgumentOutOfRangeException>(() => options.Validate());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Validate_InvalidMaxTableNestingDepth_Throws(int value)
    {
        var options = new ConversionOptions { MaxTableNestingDepth = value };
        Assert.Throws<ArgumentOutOfRangeException>(() => options.Validate());
    }

    [Theory]
    [InlineData(-1)]
    public void Validate_InvalidMaxFileSizeMB_Throws(int value)
    {
        var options = new ConversionOptions { MaxFileSizeMB = value };
        Assert.Throws<ArgumentOutOfRangeException>(() => options.Validate());
    }

    [Fact]
    public void Validate_InvalidTimeout_Throws()
    {
        var options = new ConversionOptions { Timeout = TimeSpan.Zero };
        Assert.Throws<ArgumentOutOfRangeException>(() => options.Validate());
    }

    [Fact]
    public void ValidateFileSize_WithinLimit_DoesNotThrow()
    {
        var options = new ConversionOptions { MaxFileSizeMB = 10 };
        options.ValidateFileSize(5 * 1024 * 1024); // 5MB
    }

    [Fact]
    public void ValidateFileSize_ExceedsLimit_ThrowsFileSizeLimitException()
    {
        var options = new ConversionOptions { MaxFileSizeMB = 1 };
        var exception = Assert.Throws<FileSizeLimitException>(() =>
            options.ValidateFileSize(2 * 1024 * 1024)); // 2MB

        Assert.Equal(2 * 1024 * 1024, exception.FileSize);
        Assert.Equal(1 * 1024 * 1024, exception.MaxSize);
    }

    [Fact]
    public void ValidateFileSize_ZeroLimit_DoesNotCheck()
    {
        var options = new ConversionOptions { MaxFileSizeMB = 0 };
        options.ValidateFileSize(long.MaxValue); // Should not throw
    }
}
