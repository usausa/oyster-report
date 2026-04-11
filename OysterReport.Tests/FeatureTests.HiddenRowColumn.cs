namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Fact]
    public void HiddenRowColumnShouldExcludeHiddenRow()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VisibleRow";
            sheet.Cell("A2").Value = "HiddenRow";
            sheet.Row(2).Hide();
            sheet.Cell("A3").Value = "VisibleRow2";
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HiddenRowColumnShouldExcludeHiddenRow), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("VisibleRow", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HiddenRow", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HiddenRowColumnShouldExcludeHiddenColumn()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VisibleCol";
            sheet.Cell("B1").Value = "HiddenCol";
            sheet.Column(2).Hide();
            sheet.Cell("C1").Value = "VisibleCol2";
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HiddenRowColumnShouldExcludeHiddenColumn), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("VisibleCol", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HiddenCol", text, StringComparison.Ordinal);
    }
}
