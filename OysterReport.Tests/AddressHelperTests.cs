namespace OysterReport.Tests;

public sealed class AddressHelperTests
{
    //--------------------------------------------------------------------------------
    // ToAddress
    //--------------------------------------------------------------------------------

    [Theory]
    [InlineData(1, 1, "A1")]
    [InlineData(1, 26, "Z1")]
    [InlineData(1, 27, "AA1")]
    [InlineData(1, 52, "AZ1")]
    [InlineData(1, 53, "BA1")]
    [InlineData(1, 702, "ZZ1")]
    [InlineData(1, 703, "AAA1")]
    [InlineData(100, 3, "C100")]
    [InlineData(1048576, 16384, "XFD1048576")]
    public void ToAddressShouldReturnExpectedAddress(int row, int column, string expected)
    {
        // Act
        var result = AddressHelper.ToAddress(row, column);

        // Assert
        Assert.Equal(expected, result);
    }

    //--------------------------------------------------------------------------------
    // ParseAddress
    //--------------------------------------------------------------------------------

    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("Z1", 1, 26)]
    [InlineData("AA1", 1, 27)]
    [InlineData("AZ1", 1, 52)]
    [InlineData("BA1", 1, 53)]
    [InlineData("ZZ1", 1, 702)]
    [InlineData("AAA1", 1, 703)]
    [InlineData("C100", 100, 3)]
    [InlineData("XFD1048576", 1048576, 16384)]
    public void ParseAddressShouldReturnExpectedRowAndColumn(string address, int expectedRow, int expectedColumn)
    {
        // Act
        AddressHelper.ParseAddress(address, out var row, out var column);

        // Assert
        Assert.Equal(expectedRow, row);
        Assert.Equal(expectedColumn, column);
    }

    [Fact]
    public void ParseAddressShouldBeCaseInsensitive()
    {
        // Act
        AddressHelper.ParseAddress("xfd1048576", out var row, out var column);

        // Assert
        Assert.Equal(1048576, row);
        Assert.Equal(16384, column);
    }

    [Fact]
    public void ParseAddressShouldTrimWhitespace()
    {
        // Act
        AddressHelper.ParseAddress("  B5  ", out var row, out var column);

        // Assert
        Assert.Equal(5, row);
        Assert.Equal(2, column);
    }

    [Fact]
    public void ParseAddressShouldThrowFormatExceptionWhenNoLetters()
    {
        // Act
        var exception = Record.Exception(() => AddressHelper.ParseAddress("123", out _, out _));

        // Assert
        Assert.IsType<FormatException>(exception);
    }

    [Fact]
    public void ParseAddressShouldThrowFormatExceptionWhenNoDigits()
    {
        // Act
        var exception = Record.Exception(() => AddressHelper.ParseAddress("ABC", out _, out _));

        // Assert
        Assert.IsType<FormatException>(exception);
    }

    [Fact]
    public void ParseAddressShouldThrowFormatExceptionForEmptyString()
    {
        // Act / Assert
        Assert.Throws<FormatException>(() => AddressHelper.ParseAddress(string.Empty, out _, out _));
    }

    //--------------------------------------------------------------------------------
    // Roundtrip
    //--------------------------------------------------------------------------------

    [Theory]
    [InlineData(1, 1)]
    [InlineData(1, 26)]
    [InlineData(1, 27)]
    [InlineData(100, 3)]
    [InlineData(1048576, 16384)]
    public void ToAddressThenParseAddressShouldReturnOriginalValues(int row, int column)
    {
        // Act
        var address = AddressHelper.ToAddress(row, column);
        AddressHelper.ParseAddress(address, out var parsedRow, out var parsedColumn);

        // Assert
        Assert.Equal(row, parsedRow);
        Assert.Equal(column, parsedColumn);
    }
}
