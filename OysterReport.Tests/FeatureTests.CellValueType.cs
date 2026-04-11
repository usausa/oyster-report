namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Fact]
    public void CellValueTypeShouldRenderNumericValue()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = 12345;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(CellValueTypeShouldRenderNumericValue), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("12345", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CellValueTypeShouldRenderDateValue()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = new DateTime(2025, 1, 15);
            cell.Style.DateFormat.Format = "yyyy/MM/dd";
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(CellValueTypeShouldRenderDateValue), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("2025", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CellValueTypeShouldRenderFormulaValue()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = 10;
            sheet.Cell("A2").Value = 20;
            sheet.Cell("A3").FormulaA1 = "=A1+A2";
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(CellValueTypeShouldRenderFormulaValue), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
    }
}
