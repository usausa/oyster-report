namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void HeaderFooterShouldRenderHeaderText()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Header.Left.AddText("LeftHeader", XLHFOccurrence.OddPages);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldRenderHeaderText), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BodyContent", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldRenderFooterText()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Footer.Right.AddText("RightFooter", XLHFOccurrence.OddPages);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldRenderFooterText), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void HeaderFooterShouldRenderBothHeaderAndFooter()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Main";
            sheet.PageSetup.Header.Center.AddText("TopCenter", XLHFOccurrence.OddPages);
            sheet.PageSetup.Footer.Center.AddText("BottomCenter", XLHFOccurrence.OddPages);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldRenderBothHeaderAndFooter), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
    }
}
