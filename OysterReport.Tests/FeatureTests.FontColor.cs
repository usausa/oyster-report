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

    [Fact]
    public void FontColorShouldIgnoreAlphaOfRgbColor()
    {
        // Arrange — openpyxl writes 6-digit colors as 00RRGGBB; Excel draws the text without looking at the alpha
        using var stream = TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                var zeroAlphaCell = sheet.Cell("A1");
                zeroAlphaCell.Value = "ZeroAlphaText";
                zeroAlphaCell.Style.Font.FontColor = XLColor.FromArgb(0, 0, 255);
                var halfAlphaCell = sheet.Cell("A2");
                halfAlphaCell.Value = "HalfAlphaText";
                halfAlphaCell.Style.Font.FontColor = XLColor.FromArgb(255, 0, 0);
            },
            document =>
            {
                TestWorkbookFactory.ReplaceRgbColor(document, "FF0000FF", "000000FF");
                TestWorkbookFactory.ReplaceRgbColor(document, "FFFF0000", "80FF0000");
            });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontColorShouldIgnoreAlphaOfRgbColor), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var letters = TestHelper.GetLetters(pdfBytes, 1);
        Assert.Contains(letters, static x => x.FillColor.ToRGBValues() == (0d, 0d, 1d));
        Assert.Contains(letters, static x => x.FillColor.ToRGBValues() == (1d, 0d, 0d));
        Assert.False(TestHelper.HasTransparency(pdfBytes, 1));
    }
}
