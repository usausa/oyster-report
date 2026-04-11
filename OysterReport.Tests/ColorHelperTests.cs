namespace OysterReport.Tests;

using System.Drawing;

using OysterReport.Internal;

using Xunit;

public sealed class ColorHelperTests
{
    //--------------------------------------------------------------------------------
    // NormalizeHex
    //--------------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void NormalizeHexShouldReturnTransparentBlackForNullOrWhitespace(string? input)
    {
        var result = ColorHelper.NormalizeHex(input);

        Assert.Equal("#00000000", result);
    }

    [Theory]
    [InlineData("FF0000", "#FF0000")]
    [InlineData("ff0000", "#FF0000")]
    [InlineData("FFFFFFFF", "#FFFFFFFF")]
    [InlineData("00000000", "#00000000")]
    public void NormalizeHexShouldAddHashPrefixAndUppercase(string input, string expected)
    {
        var result = ColorHelper.NormalizeHex(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("#FF0000", "#FF0000")]
    [InlineData("#ff0000", "#FF0000")]
    [InlineData("#FFFFFFFF", "#FFFFFFFF")]
    public void NormalizeHexShouldUppercaseWhenHashAlreadyPresent(string input, string expected)
    {
        var result = ColorHelper.NormalizeHex(input);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void NormalizeHexShouldTrimWhitespaceBeforeProcessing()
    {
        var result = ColorHelper.NormalizeHex("  FF0000  ");

        Assert.Equal("#FF0000", result);
    }

    //--------------------------------------------------------------------------------
    // ToHex
    //--------------------------------------------------------------------------------

    [Fact]
    public void ToHexShouldReturnCorrectHexForRed()
    {
        var result = ColorHelper.ToHex(Color.Red);

        Assert.Equal("#FFFF0000", result);
    }

    [Fact]
    public void ToHexShouldReturnCorrectHexForBlack()
    {
        var result = ColorHelper.ToHex(Color.Black);

        Assert.Equal("#FF000000", result);
    }

    [Fact]
    public void ToHexShouldReturnCorrectHexForWhite()
    {
        var result = ColorHelper.ToHex(Color.White);

        Assert.Equal("#FFFFFFFF", result);
    }

    [Fact]
    public void ToHexShouldIncludeAlphaChannel()
    {
        var color = Color.FromArgb(128, 0, 255, 0);

        var result = ColorHelper.ToHex(color);

        Assert.Equal("#8000FF00", result);
    }

    //--------------------------------------------------------------------------------
    // ApplyTint
    //--------------------------------------------------------------------------------

    [Fact]
    public void ApplyTintShouldReturnSameColorWhenTintIsNaN()
    {
        var result = ColorHelper.ApplyTint(Color.Red, double.NaN);

        Assert.Equal(Color.Red, result);
    }

    [Fact]
    public void ApplyTintShouldReturnSameColorWhenTintIsZero()
    {
        var result = ColorHelper.ApplyTint(Color.Red, 0d);

        Assert.Equal(Color.Red, result);
    }

    [Fact]
    public void ApplyTintShouldReturnBlackWhenTintIsNegativeOneOnGray()
    {
        var gray = Color.FromArgb(255, 128, 128, 128);

        var result = ColorHelper.ApplyTint(gray, -1d);

        Assert.Equal(0, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(0, result.B);
    }

    [Fact]
    public void ApplyTintShouldReturnWhiteWhenTintIsOneOnGray()
    {
        var gray = Color.FromArgb(255, 128, 128, 128);

        var result = ColorHelper.ApplyTint(gray, 1d);

        Assert.Equal(255, result.R);
        Assert.Equal(255, result.G);
        Assert.Equal(255, result.B);
    }

    [Fact]
    public void ApplyTintShouldDarkenColorWhenTintIsNegative()
    {
        var result = ColorHelper.ApplyTint(Color.Red, -0.5d);

        Assert.True(result.R < Color.Red.R);
        Assert.Equal(255, result.A);
    }

    [Fact]
    public void ApplyTintShouldLightenColorWhenTintIsPositive()
    {
        var result = ColorHelper.ApplyTint(Color.Red, 0.5d);

        Assert.True(result.G > Color.Red.G);
        Assert.Equal(255, result.A);
    }

    [Fact]
    public void ApplyTintShouldPreserveAlphaChannel()
    {
        var color = Color.FromArgb(100, 200, 100, 50);

        var result = ColorHelper.ApplyTint(color, 0.5d);

        Assert.Equal(100, result.A);
    }
}
