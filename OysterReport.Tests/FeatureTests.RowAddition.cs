namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Fact]
    public void RowAdditionShouldContainRowsFromInsertCopyBelow()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
            sheet.Cell("A3").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Item");
        var row1 = template.InsertCopyBelow();
        var row2 = row1.InsertCopyBelow();
        var row3 = row2.InsertCopyBelow();
        row1.ReplacePlaceholder("Item", "ItemA");
        row2.ReplacePlaceholder("Item", "ItemB");
        row3.ReplacePlaceholder("Item", "ItemC");
        template.Delete();

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromInsertCopyBelow),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("ItemA", text, StringComparison.Ordinal);
        Assert.Contains("ItemB", text, StringComparison.Ordinal);
        Assert.Contains("ItemC", text, StringComparison.Ordinal);
        Assert.Contains("Footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldContainRowsFromInsertCopyAfter()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Row}}";
            sheet.Cell("A3").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var templateRow = sheet.FindRow("Row");
        var last = templateRow;
        foreach (var label in new[] { "Row1", "Row2", "Row3" })
        {
            last = templateRow.InsertCopyAfter(last);
            last.ReplacePlaceholder("Row", label);
        }

        templateRow.Delete();

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromInsertCopyAfter),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Row1", text, StringComparison.Ordinal);
        Assert.Contains("Row2", text, StringComparison.Ordinal);
        Assert.Contains("Row3", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldPreserveStyleAfterCopy()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var templateCell = sheet.Cell("A1");
            templateCell.Value = "{{StyledItem}}";
            templateCell.Style.Font.Bold = true;
            templateCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 230, 255);
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("StyledItem");
        var copy = template.InsertCopyBelow();
        copy.ReplacePlaceholder("StyledItem", "CopiedRow");
        template.Delete();

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(RowAdditionShouldPreserveStyleAfterCopy), workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("CopiedRow", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldHandleZeroDetailRows()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
            sheet.Cell("A3").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.FindRow("Item").Delete();

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(RowAdditionShouldHandleZeroDetailRows), workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("Footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldContainRowsFromMultiRowRangeExpansion()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Name}}";
            sheet.Cell("A3").Value = "{{Detail}}";
            sheet.Cell("A4").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var templateRange = sheet.FindRows("Name");
        var last = templateRange;
        foreach (var (name, detail) in new[] { ("Alice", "Detail1"), ("Bob", "Detail2") })
        {
            last = templateRange.InsertCopyAfter(last);
            last.ReplacePlaceholder("Name", name);
            last.ReplacePlaceholder("Detail", detail);
        }

        sheet.DeleteRows(templateRange.StartRow, templateRange.EndRow);

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromMultiRowRangeExpansion),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alice", text, StringComparison.Ordinal);
        Assert.Contains("Detail1", text, StringComparison.Ordinal);
        Assert.Contains("Bob", text, StringComparison.Ordinal);
        Assert.Contains("Detail2", text, StringComparison.Ordinal);
    }
}
