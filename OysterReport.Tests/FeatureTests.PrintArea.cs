namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Fact]
    public void PrintAreaShouldExcludeContentOutsidePrintArea()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "InPrintArea";
            sheet.Cell("A2").Value = "InPrintArea2";
            sheet.Cell("A10").Value = "OutsidePrintArea";
            sheet.PageSetup.PrintAreas.Add("A1:B5");
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(PrintAreaShouldExcludeContentOutsidePrintArea), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("InPrintArea", text, StringComparison.Ordinal);
        Assert.DoesNotContain("OutsidePrintArea", text, StringComparison.Ordinal);
    }
}
