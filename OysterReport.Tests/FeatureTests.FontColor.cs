namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Theory]
    [InlineData("RedText", 255, 0, 0)]
    [InlineData("BlueText", 0, 0, 255)]
    [InlineData("GreenText", 0, 128, 0)]
    public void FontColorShouldRenderColoredText(string cellValue, int r, int g, int b)
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Font.FontColor = XLColor.FromArgb(r, g, b);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(FontColorShouldRenderColoredText)}_{cellValue}",
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontColorShouldRenderThemeColorText()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeColorText";
            cell.Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.4);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontColorShouldRenderThemeColorText), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeColorText", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
