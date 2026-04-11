namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Theory]
    [InlineData(XLBorderStyleValues.Thin, "ThinBorder")]
    [InlineData(XLBorderStyleValues.Medium, "MediumBorder")]
    [InlineData(XLBorderStyleValues.Thick, "ThickBorder")]
    [InlineData(XLBorderStyleValues.Double, "DoubleBorder")]
    [InlineData(XLBorderStyleValues.Dashed, "DashedBorder")]
    [InlineData(XLBorderStyleValues.Dotted, "DottedBorder")]
    public void BorderShouldRenderCellWithBorder(XLBorderStyleValues borderStyle, string cellValue)
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = cellValue;
            cell.Style.Border.OutsideBorder = borderStyle;
            cell.Style.Border.OutsideBorderColor = XLColor.Black;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(BorderShouldRenderCellWithBorder)}_{borderStyle}",
            stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void BorderShouldRenderColoredBorder()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = "ColoredBorder";
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            cell.Style.Border.OutsideBorderColor = XLColor.FromArgb(255, 0, 0);
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(BorderShouldRenderColoredBorder), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ColoredBorder", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void BorderShouldRenderTableWithAllSideBorders()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            for (var row = 1; row <= 3; row++)
            {
                for (var col = 1; col <= 3; col++)
                {
                    var cell = sheet.Cell(row, col);
                    cell.Value = $"R{row}C{col}";
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(BorderShouldRenderTableWithAllSideBorders), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("R1C1", text, StringComparison.Ordinal);
        Assert.Contains("R3C3", text, StringComparison.Ordinal);
    }
}
