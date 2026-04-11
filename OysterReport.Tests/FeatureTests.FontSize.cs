namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Theory]
    [InlineData(6d, "TinyText")]
    [InlineData(11d, "NormalText")]
    [InlineData(18d, "LargeText")]
    [InlineData(24d, "HugeText")]
    public void FontSizeShouldRenderVariousSizes(double fontSize, string cellValue)
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Font.FontSize = fontSize;
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(FontSizeShouldRenderVariousSizes)}_{fontSize}pt",
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontSizeShouldRenderMultipleSizesOnOnePage()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Size8";
            sheet.Cell("A1").Style.Font.FontSize = 8d;
            sheet.Cell("A2").Value = "Size12";
            sheet.Cell("A2").Style.Font.FontSize = 12d;
            sheet.Cell("A3").Value = "Size16";
            sheet.Cell("A3").Style.Font.FontSize = 16d;
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(FontSizeShouldRenderMultipleSizesOnOnePage),
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Size8", text, StringComparison.Ordinal);
        Assert.Contains("Size12", text, StringComparison.Ordinal);
        Assert.Contains("Size16", text, StringComparison.Ordinal);
    }
}
