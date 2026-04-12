namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void MultiSheetShouldRenderOnePagePerSheet()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Sheet1").Cell("A1").Value = "ContentSheet1";
            workbook.AddWorksheet("Sheet2").Cell("A1").Value = "ContentSheet2";
            workbook.AddWorksheet("Sheet3").Cell("A1").Value = "ContentSheet3";
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(MultiSheetShouldRenderOnePagePerSheet), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.True(TestHelper.GetPageCount(pdfBytes) >= 3);
    }

    [Fact]
    public void MultiSheetShouldContainTextFromAllSheets()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Alpha").Cell("A1").Value = "AlphaSheet";
            workbook.AddWorksheet("Beta").Cell("A1").Value = "BetaSheet";
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(MultiSheetShouldContainTextFromAllSheets), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("AlphaSheet", text, StringComparison.Ordinal);
        Assert.Contains("BetaSheet", text, StringComparison.Ordinal);
    }
}
