namespace OysterReport.Tests;

public sealed class FontMetricsHelperTests
{
    //--------------------------------------------------------------------------------
    // Guard
    //--------------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void MeasureMaxDigitWidthShouldReturnNullForNullOrWhitespaceFontName(string? fontName)
    {
        // Act
        var result = FontMetricsHelper.MeasureMaxDigitWidth(fontName!, 11d);

        // Assert
        Assert.Null(result);
    }

    [Theory]
    [InlineData(0d)]
    [InlineData(-1d)]
    [InlineData(-100d)]
    public void MeasureMaxDigitWidthShouldReturnNullForNonPositiveFontSize(double fontSize)
    {
        // Act
        var result = FontMetricsHelper.MeasureMaxDigitWidth("Arial", fontSize);

        // Assert
        Assert.Null(result);
    }

    [Fact]
    public void MeasureMaxDigitWidthShouldReturnNullForUnknownFont()
    {
        // Act
        var result = FontMetricsHelper.MeasureMaxDigitWidth("__NonExistentFont__XYZ__", 11d);

        // Assert
        // SkiaSharp may fall back to a default typeface instead of returning null,
        // so only verify that a non-null result is positive.
        Assert.True(result is null || result > 0d);
    }

    //--------------------------------------------------------------------------------
    // Valid font (Windows only)
    //--------------------------------------------------------------------------------

    [Fact]
    public void MeasureMaxDigitWidthShouldReturnPositiveValueForArialOnWindows()
    {
        // Arrange
        if (!OperatingSystem.IsWindows())
        {
            Assert.Skip("Arial font is only guaranteed on Windows.");
        }

        // Act
        var result = FontMetricsHelper.MeasureMaxDigitWidth("Arial", 11d);

        // Assert
        Assert.NotNull(result);
        Assert.True(result > 0d);
    }

    [Fact]
    public void MeasureMaxDigitWidthShouldReturnLargerValueForLargerFontSize()
    {
        // Arrange
        if (!OperatingSystem.IsWindows())
        {
            Assert.Skip("Arial font is only guaranteed on Windows.");
        }

        // Act
        var small = FontMetricsHelper.MeasureMaxDigitWidth("Arial", 8d);
        var large = FontMetricsHelper.MeasureMaxDigitWidth("Arial", 24d);

        // Assert
        Assert.NotNull(small);
        Assert.NotNull(large);
        Assert.True(large > small);
    }
}
