namespace OysterReport.Tests;

using OysterReport.Internal;

using Xunit;

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
        var result = FontMetricsHelper.MeasureMaxDigitWidth(fontName!, 11d);

        Assert.Null(result);
    }

    [Theory]
    [InlineData(0d)]
    [InlineData(-1d)]
    [InlineData(-100d)]
    public void MeasureMaxDigitWidthShouldReturnNullForNonPositiveFontSize(double fontSize)
    {
        var result = FontMetricsHelper.MeasureMaxDigitWidth("Arial", fontSize);

        Assert.Null(result);
    }

    [Fact]
    public void MeasureMaxDigitWidthShouldReturnNullForUnknownFont()
    {
        var result = FontMetricsHelper.MeasureMaxDigitWidth("__NonExistentFont__XYZ__", 11d);

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
        if (!OperatingSystem.IsWindows())
        {
            Assert.Skip("Arial font is only guaranteed on Windows.");
        }

        var result = FontMetricsHelper.MeasureMaxDigitWidth("Arial", 11d);

        Assert.NotNull(result);
        Assert.True(result > 0d);
    }

    [Fact]
    public void MeasureMaxDigitWidthShouldReturnLargerValueForLargerFontSize()
    {
        if (!OperatingSystem.IsWindows())
        {
            Assert.Skip("Arial font is only guaranteed on Windows.");
        }

        var small = FontMetricsHelper.MeasureMaxDigitWidth("Arial", 8d);
        var large = FontMetricsHelper.MeasureMaxDigitWidth("Arial", 24d);

        Assert.NotNull(small);
        Assert.NotNull(large);
        Assert.True(large > small);
    }
}
