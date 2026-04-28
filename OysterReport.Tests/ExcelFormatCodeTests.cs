namespace OysterReport.Tests;

using OysterReport.Internal.OpenXml;

public sealed class ExcelFormatCodeTests
{
    //--------------------------------------------------------------------------------
    // IsDateTime
    //--------------------------------------------------------------------------------

    [Theory]
    [InlineData("yyyy-mm-dd")]
    [InlineData("yyyy/mm/dd hh:mm")]
    [InlineData("h:mm AM/PM")]
    [InlineData("[h]:mm:ss")]
    [InlineData("[mm]:ss")]
    [InlineData("d-mmm-yy")]
    [InlineData("mmss.0")]
    public void IsDateTimeShouldReturnTrueForDateOrTimeFormats(string code)
    {
        Assert.True(ExcelFormatCode.IsDateTime(code));
    }

    [Theory]
    [InlineData("")]
    [InlineData("General")]
    [InlineData("@")]
    [InlineData("#,##0")]
    [InlineData("0.00")]
    [InlineData("0.0%")]
    [InlineData("0.00E+00")]
    [InlineData("\"$\"#,##0")]
    [InlineData("[Red]#,##0")]
    [InlineData("0;-0;\"zero\"")]
    public void IsDateTimeShouldReturnFalseForNumericOrTextFormats(string code)
    {
        Assert.False(ExcelFormatCode.IsDateTime(code));
    }

    //--------------------------------------------------------------------------------
    // Format - Numeric
    //--------------------------------------------------------------------------------

    [Fact]
    public void FormatShouldRenderThousandsSeparator()
    {
        Assert.Equal("1,234,567", ExcelFormatCode.Format(1234567, "#,##0"));
    }

    [Fact]
    public void FormatShouldRenderFixedDecimals()
    {
        Assert.Equal("3.14", ExcelFormatCode.Format(3.14159, "0.00"));
    }

    [Fact]
    public void FormatShouldRenderPercentage()
    {
        Assert.Equal("12.3%", ExcelFormatCode.Format(0.123, "0.0%"));
    }

    [Fact]
    public void FormatShouldRenderCurrencyLiteralPrefix()
    {
        Assert.Equal("$1,500", ExcelFormatCode.Format(1500, "\"$\"#,##0"));
    }

    [Fact]
    public void FormatShouldRenderNegativeSection()
    {
        Assert.Equal("(1,234)", ExcelFormatCode.Format(-1234, "#,##0;(#,##0)"));
    }

    [Fact]
    public void FormatShouldRenderZeroSection()
    {
        Assert.Equal("zero", ExcelFormatCode.Format(0, "0;-0;\"zero\""));
    }

    [Fact]
    public void FormatShouldRenderTrailingCommaScaling()
    {
        Assert.Equal("1,500", ExcelFormatCode.Format(1500000, "#,##0,"));
    }

    [Fact]
    public void FormatShouldRenderScientificNotation()
    {
        Assert.Equal("1.23E+04", ExcelFormatCode.Format(12345d, "0.00E+00"));
    }

    [Fact]
    public void FormatShouldStripColorBracketDirective()
    {
        var result = ExcelFormatCode.Format(250, "[Red]#,##0");
        Assert.Equal("250", result);
        Assert.DoesNotContain("Red", result, StringComparison.Ordinal);
        Assert.DoesNotContain("[", result, StringComparison.Ordinal);
    }

    [Fact]
    public void FormatShouldFallBackToGeneralForFractionPattern()
    {
        Assert.Equal("1.5", ExcelFormatCode.Format(1.5, "# ?/?"));
    }

    [Fact]
    public void FormatShouldRenderGeneralForGeneralFormatString()
    {
        Assert.Equal("42.5", ExcelFormatCode.Format(42.5, "General"));
    }

    [Fact]
    public void FormatShouldRenderGeneralForEmptyFormatString()
    {
        Assert.Equal("42.5", ExcelFormatCode.Format(42.5, string.Empty));
    }

    [Fact]
    public void FormatShouldHandleSinglePositiveSection()
    {
        Assert.Equal("9,876,543", ExcelFormatCode.Format(9876543, "#,##0"));
    }

    [Fact]
    public void FormatShouldHandleSimplePercent()
    {
        Assert.Equal("50%", ExcelFormatCode.Format(0.5, "0%"));
    }

    //--------------------------------------------------------------------------------
    // Format - Date / Time
    //--------------------------------------------------------------------------------

    [Fact]
    public void FormatShouldRenderCustomDateFormat()
    {
        var oa = new DateTime(2025, 1, 15).ToOADate();
        Assert.Equal("2025-01-15", ExcelFormatCode.Format(oa, "yyyy-mm-dd"));
    }

    [Fact]
    public void FormatShouldDistinguishMonthAndMinute()
    {
        var oa = new DateTime(2025, 3, 8, 9, 5, 0).ToOADate();
        Assert.Equal("2025/03/08 09:05", ExcelFormatCode.Format(oa, "yyyy/mm/dd hh:mm"));
    }

    [Fact]
    public void FormatShouldRenderAmPmTime()
    {
        var oa = new DateTime(2025, 1, 1, 14, 30, 0).ToOADate();
        Assert.Equal("2:30 PM", ExcelFormatCode.Format(oa, "h:mm AM/PM"));
    }

    [Fact]
    public void FormatShouldRenderElapsedHours()
    {
        Assert.Equal("25:30:00", ExcelFormatCode.Format(1.0625, "[h]:mm:ss"));
    }

    [Fact]
    public void FormatShouldRenderTwoDigitYear()
    {
        var oa = new DateTime(2025, 6, 1).ToOADate();
        Assert.Equal("25-06-01", ExcelFormatCode.Format(oa, "yy-mm-dd"));
    }

    [Fact]
    public void FormatShouldRenderAbbreviatedMonthName()
    {
        var oa = new DateTime(2025, 3, 1).ToOADate();
        Assert.Equal("Mar", ExcelFormatCode.Format(oa, "mmm"));
    }

    [Fact]
    public void FormatShouldRenderFullMonthName()
    {
        var oa = new DateTime(2025, 3, 1).ToOADate();
        Assert.Equal("March", ExcelFormatCode.Format(oa, "mmmm"));
    }

    [Fact]
    public void FormatShouldRenderTwelveHourMidnight()
    {
        var oa = new DateTime(2025, 1, 1, 0, 0, 0).ToOADate();
        Assert.Equal("12:00 AM", ExcelFormatCode.Format(oa, "h:mm AM/PM"));
    }
}
