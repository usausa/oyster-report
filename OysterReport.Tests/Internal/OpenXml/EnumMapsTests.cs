namespace OysterReport.Tests.Internal.OpenXml;

public sealed class EnumMapsTests
{
    //--------------------------------------------------------------------------------
    // ToBorderStyle
    //--------------------------------------------------------------------------------

    [Fact]
    public void ToBorderStyleShouldReturnNoneWhenNull()
    {
        // Act
        var result = EnumMaps.ToBorderStyle(null);

        // Assert
        Assert.Equal(BorderLineStyle.None, result);
    }

    [Theory]
    [InlineData("Thin", nameof(BorderLineStyle.Thin))]
    [InlineData("Medium", nameof(BorderLineStyle.Medium))]
    [InlineData("Thick", nameof(BorderLineStyle.Thick))]
    [InlineData("Double", nameof(BorderLineStyle.Double))]
    [InlineData("Hair", nameof(BorderLineStyle.Hair))]
    [InlineData("Dotted", nameof(BorderLineStyle.Dotted))]
    [InlineData("Dashed", nameof(BorderLineStyle.Dashed))]
    [InlineData("DashDot", nameof(BorderLineStyle.DashDot))]
    [InlineData("DashDotDot", nameof(BorderLineStyle.DashDotDot))]
    [InlineData("MediumDashed", nameof(BorderLineStyle.MediumDashed))]
    [InlineData("MediumDashDot", nameof(BorderLineStyle.MediumDashDot))]
    [InlineData("MediumDashDotDot", nameof(BorderLineStyle.MediumDashDotDot))]
    [InlineData("SlantDashDot", nameof(BorderLineStyle.SlantDashDot))]
    public void ToBorderStyleShouldMapKnownValues(string sourceName, string expectedName)
    {
        // Arrange
        var enumValue = (BorderStyleValues)typeof(BorderStyleValues).GetProperty(sourceName)!.GetValue(null)!;
        var expected = Enum.Parse<BorderLineStyle>(expectedName);

        // Act
        var result = EnumMaps.ToBorderStyle(enumValue);

        // Assert
        Assert.Equal(expected, result);
    }

    //--------------------------------------------------------------------------------
    // ToHorizontalAlignment
    //--------------------------------------------------------------------------------

    [Fact]
    public void ToHorizontalAlignmentShouldReturnGeneralWhenNull()
    {
        // Act
        var result = EnumMaps.ToHorizontalAlignment(null);

        // Assert
        Assert.Equal(HorizontalAlignment.General, result);
    }

    [Theory]
    [InlineData("Left", nameof(HorizontalAlignment.Left))]
    [InlineData("Center", nameof(HorizontalAlignment.Center))]
    [InlineData("Right", nameof(HorizontalAlignment.Right))]
    [InlineData("Justify", nameof(HorizontalAlignment.Justify))]
    [InlineData("CenterContinuous", nameof(HorizontalAlignment.CenterContinuous))]
    [InlineData("Distributed", nameof(HorizontalAlignment.Distributed))]
    [InlineData("Fill", nameof(HorizontalAlignment.Fill))]
    public void ToHorizontalAlignmentShouldMapKnownValues(string sourceName, string expectedName)
    {
        // Arrange
        var enumValue = (HorizontalAlignmentValues)typeof(HorizontalAlignmentValues).GetProperty(sourceName)!.GetValue(null)!;
        var expected = Enum.Parse<HorizontalAlignment>(expectedName);

        // Act
        var result = EnumMaps.ToHorizontalAlignment(enumValue);

        // Assert
        Assert.Equal(expected, result);
    }

    //--------------------------------------------------------------------------------
    // ToVerticalAlignment
    //--------------------------------------------------------------------------------

    [Fact]
    public void ToVerticalAlignmentShouldReturnBottomWhenNull()
    {
        // Act
        var result = EnumMaps.ToVerticalAlignment(null);

        // Assert
        Assert.Equal(VerticalAlignment.Bottom, result);
    }

    [Theory]
    [InlineData("Top", nameof(VerticalAlignment.Top))]
    [InlineData("Center", nameof(VerticalAlignment.Center))]
    [InlineData("Justify", nameof(VerticalAlignment.Justify))]
    [InlineData("Distributed", nameof(VerticalAlignment.Distributed))]
    public void ToVerticalAlignmentShouldMapKnownValues(string sourceName, string expectedName)
    {
        // Arrange
        var enumValue = (VerticalAlignmentValues)typeof(VerticalAlignmentValues).GetProperty(sourceName)!.GetValue(null)!;
        var expected = Enum.Parse<VerticalAlignment>(expectedName);

        // Act
        var result = EnumMaps.ToVerticalAlignment(enumValue);

        // Assert
        Assert.Equal(expected, result);
    }

    //--------------------------------------------------------------------------------
    // ToPaperSize
    //--------------------------------------------------------------------------------

    [Fact]
    public void ToPaperSizeShouldReturnA4PaperWhenNull()
    {
        // Act
        var result = EnumMaps.ToPaperSize(null);

        // Assert
        Assert.Equal(PaperSize.A4Paper, result);
    }

    [Theory]
    [InlineData(1u, PaperSize.LetterPaper)]
    [InlineData(8u, PaperSize.A3Paper)]
    [InlineData(9u, PaperSize.A4Paper)]
    [InlineData(11u, PaperSize.A5Paper)]
    public void ToPaperSizeShouldReturnExpectedKnownPaperSize(uint code, PaperSize expected)
    {
        // Act
        var result = EnumMaps.ToPaperSize(code);

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ToPaperSizeShouldReturnDefaultForUnknownCode()
    {
        // Act
        var result = EnumMaps.ToPaperSize(9999u);

        // Assert
        Assert.Equal(PaperSize.Default, result);
    }

    //--------------------------------------------------------------------------------
    // ToPageOrientation
    //--------------------------------------------------------------------------------

    [Fact]
    public void ToPageOrientationShouldReturnDefaultWhenNull()
    {
        // Act
        var result = EnumMaps.ToPageOrientation(null);

        // Assert
        Assert.Equal(PageOrientation.Default, result);
    }

    [Fact]
    public void ToPageOrientationShouldReturnPortrait()
    {
        // Act
        var result = EnumMaps.ToPageOrientation(OrientationValues.Portrait);

        // Assert
        Assert.Equal(PageOrientation.Portrait, result);
    }

    [Fact]
    public void ToPageOrientationShouldReturnLandscape()
    {
        // Act
        var result = EnumMaps.ToPageOrientation(OrientationValues.Landscape);

        // Assert
        Assert.Equal(PageOrientation.Landscape, result);
    }

    [Fact]
    public void ToPageOrientationShouldReturnDefaultForOtherValues()
    {
        // Act
        var result = EnumMaps.ToPageOrientation(OrientationValues.Default);

        // Assert
        Assert.Equal(PageOrientation.Default, result);
    }

    //--------------------------------------------------------------------------------
    // ToFillPattern
    //--------------------------------------------------------------------------------

    [Fact]
    public void ToFillPatternShouldReturnNoneWhenNull()
    {
        // Act
        var result = EnumMaps.ToFillPattern(null);

        // Assert
        Assert.Equal(FillPattern.None, result);
    }

    [Theory]
    [InlineData("None", nameof(FillPattern.None))]
    [InlineData("Solid", nameof(FillPattern.Solid))]
    [InlineData("Gray125", nameof(FillPattern.Gray125))]
    [InlineData("Gray0625", nameof(FillPattern.Gray0625))]
    [InlineData("DarkGray", nameof(FillPattern.DarkGray))]
    [InlineData("MediumGray", nameof(FillPattern.MediumGray))]
    [InlineData("LightGray", nameof(FillPattern.LightGray))]
    [InlineData("DarkHorizontal", nameof(FillPattern.DarkHorizontal))]
    [InlineData("DarkVertical", nameof(FillPattern.DarkVertical))]
    [InlineData("DarkDown", nameof(FillPattern.DarkDown))]
    [InlineData("DarkUp", nameof(FillPattern.DarkUp))]
    [InlineData("DarkGrid", nameof(FillPattern.DarkGrid))]
    [InlineData("DarkTrellis", nameof(FillPattern.DarkTrellis))]
    [InlineData("LightHorizontal", nameof(FillPattern.LightHorizontal))]
    [InlineData("LightVertical", nameof(FillPattern.LightVertical))]
    [InlineData("LightDown", nameof(FillPattern.LightDown))]
    [InlineData("LightUp", nameof(FillPattern.LightUp))]
    [InlineData("LightGrid", nameof(FillPattern.LightGrid))]
    [InlineData("LightTrellis", nameof(FillPattern.LightTrellis))]
    public void ToFillPatternShouldMapKnownValues(string sourceName, string expectedName)
    {
        // Arrange
        var enumValue = (PatternValues)typeof(PatternValues).GetProperty(sourceName)!.GetValue(null)!;
        var expected = Enum.Parse<FillPattern>(expectedName);

        // Act
        var result = EnumMaps.ToFillPattern(enumValue);

        // Assert
        Assert.Equal(expected, result);
    }
}
