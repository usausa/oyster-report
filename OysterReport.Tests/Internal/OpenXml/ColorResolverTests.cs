namespace OysterReport.Tests.Internal.OpenXml;

public sealed class ColorResolverTests
{
    private static readonly ArgbColor[] Theme =
    [
        ArgbColor.White,
        ArgbColor.Black,
        new(0xFF, 0xEE, 0xEC, 0xE1),
        new(0xFF, 0x1F, 0x49, 0x7D),
        new(0xFF, 0x4F, 0x81, 0xBD),
        new(0xFF, 0xC0, 0x50, 0x4D)
    ];

    //--------------------------------------------------------------------------------
    // TryGetThemeColor
    //--------------------------------------------------------------------------------

    [Fact]
    public void TryGetThemeColorShouldReturnTrueForValidIndex()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);

        // Act
        var ok = resolver.TryGetThemeColor(4, out var color);

        // Assert
        Assert.True(ok);
        Assert.Equal(Theme[4], color);
    }

    [Fact]
    public void TryGetThemeColorShouldReturnFalseForNegativeIndex()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);

        // Act
        var ok = resolver.TryGetThemeColor(-1, out _);

        // Assert
        Assert.False(ok);
    }

    [Fact]
    public void TryGetThemeColorShouldReturnFalseForIndexBeyondPalette()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);

        // Act
        var ok = resolver.TryGetThemeColor(Theme.Length, out _);

        // Assert
        Assert.False(ok);
    }

    //--------------------------------------------------------------------------------
    // Resolve – fallback paths
    //--------------------------------------------------------------------------------

    [Fact]
    public void ResolveShouldReturnNormalizedFallbackWhenColorIsNull()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);

        // Act
        var result = resolver.Resolve(null, "ff0000");

        // Assert
        Assert.Equal("#FF0000", result);
    }

    [Fact]
    public void ResolveShouldReturnFallbackWhenAutoIsTrue()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Auto = true };

        // Act
        var result = resolver.Resolve(color, "#abcdef");

        // Assert
        Assert.Equal("#ABCDEF", result);
    }

    [Fact]
    public void ResolveShouldReturnFallbackWhenColorHasNoSpecifier()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor();

        // Act
        var result = resolver.Resolve(color, "#FF112233");

        // Assert
        Assert.Equal("#FF112233", result);
    }

    //--------------------------------------------------------------------------------
    // Resolve – Rgb
    //--------------------------------------------------------------------------------

    [Fact]
    public void ResolveShouldReturnRgbValueWhenSpecified()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Rgb = "FF1234AB" };

        // Act
        var result = resolver.Resolve(color, "#000000");

        // Assert
        Assert.Equal("#FF1234AB", result);
    }

    [Fact]
    public void ResolveShouldPrependAlphaWhenRgbIsSixCharacters()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Rgb = "1234AB" };

        // Act
        var result = resolver.Resolve(color, "#000000");

        // Assert
        Assert.Equal("#FF1234AB", result);
    }

    [Fact]
    public void ResolveShouldUppercaseLowercaseRgbValue()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Rgb = "ff1234ab" };

        // Act
        var result = resolver.Resolve(color, "#000000");

        // Assert
        Assert.Equal("#FF1234AB", result);
    }

    //--------------------------------------------------------------------------------
    // Resolve – Theme
    //--------------------------------------------------------------------------------

    [Fact]
    public void ResolveShouldReturnThemeColorWhenIndexIsValid()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Theme = 4u };

        // Act
        var result = resolver.Resolve(color, "#000000");

        // Assert
        Assert.Equal("#FF4F81BD", result);
    }

    [Fact]
    public void ResolveShouldFallbackWhenThemeIndexOutOfRange()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Theme = 999u };

        // Act
        var result = resolver.Resolve(color, "#FF112233");

        // Assert
        Assert.Equal("#FF112233", result);
    }

    [Fact]
    public void ResolveShouldApplyTintWhenThemeColorHasTint()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var withoutTint = new TabColor { Theme = 4u };
        var withTint = new TabColor { Theme = 4u, Tint = 0.5d };

        // Act
        var baseHex = resolver.Resolve(withoutTint, "#000000");
        var tintedHex = resolver.Resolve(withTint, "#000000");

        // Assert
        Assert.NotEqual(baseHex, tintedHex);
        Assert.StartsWith("#FF", tintedHex, StringComparison.Ordinal);
    }

    [Fact]
    public void ResolveShouldNotApplyTintWhenTintIsZero()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Theme = 4u, Tint = 0d };

        // Act
        var result = resolver.Resolve(color, "#000000");

        // Assert
        Assert.Equal("#FF4F81BD", result);
    }

    //--------------------------------------------------------------------------------
    // Resolve – Indexed
    //--------------------------------------------------------------------------------

    [Fact]
    public void ResolveShouldReturnIndexedPaletteValue()
    {
        // Arrange — index 2 in the standard palette is FF0000 (red).
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Indexed = 2u };

        // Act
        var result = resolver.Resolve(color, "#000000");

        // Assert
        Assert.Equal("#FFFF0000", result);
    }

    [Fact]
    public void ResolveShouldReturnFallbackWhenIndexedOutOfRange()
    {
        // Arrange
        var resolver = new ColorResolver(Theme);
        var color = new TabColor { Indexed = 9999u };

        // Act
        var result = resolver.Resolve(color, "#FF112233");

        // Assert
        Assert.Equal("#FF112233", result);
    }
}
