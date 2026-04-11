namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Theory]
    [InlineData("YellowBg", 255, 255, 0)]
    [InlineData("LightBlueBg", 173, 216, 230)]
    [InlineData("GrayBg", 192, 192, 192)]
    public void FillColorShouldRenderTextOnColoredBackground(string cellValue, int r, int g, int b)
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(r, g, b);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(FillColorShouldRenderTextOnColoredBackground)}_{cellValue}",
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FillColorShouldRenderMultipleDifferentBackgroundColors()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var redCell = sheet.Cell("A1");
            redCell.Value = "RedBg";
            redCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 200, 200);
            var greenCell = sheet.Cell("A2");
            greenCell.Value = "GreenBg";
            greenCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 255, 200);
            var blueCell = sheet.Cell("A3");
            blueCell.Value = "BlueBg";
            blueCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 200, 255);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(FillColorShouldRenderMultipleDifferentBackgroundColors),
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("RedBg", text, StringComparison.Ordinal);
        Assert.Contains("GreenBg", text, StringComparison.Ordinal);
        Assert.Contains("BlueBg", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FillColorShouldRenderThemeBackgroundColor()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeBg";
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.25);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FillColorShouldRenderThemeBackgroundColor), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeBg", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
