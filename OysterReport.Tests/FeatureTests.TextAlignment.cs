namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Theory]
    [InlineData(XLAlignmentHorizontalValues.Left, "LeftAligned")]
    [InlineData(XLAlignmentHorizontalValues.Center, "CenterAligned")]
    [InlineData(XLAlignmentHorizontalValues.Right, "RightAligned")]
    public void TextAlignmentShouldRenderHorizontalAlignment(XLAlignmentHorizontalValues alignment, string cellValue)
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Alignment.Horizontal = alignment;
            sheet.Column(1).Width = 30d;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(TextAlignmentShouldRenderHorizontalAlignment)}_{alignment}",
            stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(XLAlignmentVerticalValues.Top, "TopAligned")]
    [InlineData(XLAlignmentVerticalValues.Center, "MiddleAligned")]
    [InlineData(XLAlignmentVerticalValues.Bottom, "BottomAligned")]
    public void TextAlignmentShouldRenderVerticalAlignment(XLAlignmentVerticalValues alignment, string cellValue)
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Alignment.Vertical = alignment;
            sheet.Row(1).Height = 40d;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(TextAlignmentShouldRenderVerticalAlignment)}_{alignment}",
            stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void TextAlignmentShouldRenderWrappedText()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "This is a long text that should wrap inside the cell boundaries";
            cell.Style.Alignment.WrapText = true;
            sheet.Column(1).Width = 20d;
            sheet.Row(1).Height = 50d;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(TextAlignmentShouldRenderWrappedText), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
    }
}
