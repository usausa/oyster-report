namespace OysterReport.Tests;

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

    [Fact]
    public void PrintAreaShouldContainRowsInsertedBelowLastRow()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
            sheet.Cell("A5").Value = "OutsidePrintArea";
            sheet.PageSetup.PrintAreas.Add("A1:B2");
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var items = new[] { "ItemA", "ItemB", "ItemC" };
        var template = sheet.FindRow("Item");
        for (var i = 1; i < items.Length; i++)
        {
            template.InsertCopyBelow();
        }

        sheet.ReplacePlaceholders(items.Select(static x => new Dictionary<string, string?> { ["Item"] = x }));

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(PrintAreaShouldContainRowsInsertedBelowLastRow), workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("ItemA", text, StringComparison.Ordinal);
        Assert.Contains("ItemB", text, StringComparison.Ordinal);
        Assert.Contains("ItemC", text, StringComparison.Ordinal);
        Assert.DoesNotContain("OutsidePrintArea", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PrintAreaShouldFollowRowsShiftedUpByDeletionAboveStart()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "AbovePrintArea";
            sheet.Cell("A2").Value = "Deleted1";
            sheet.Cell("A3").Value = "Deleted2";
            sheet.Cell("A4").Value = "Kept1";
            sheet.Cell("A5").Value = "Kept2";
            sheet.Cell("A6").Value = "BelowPrintArea";
            sheet.PageSetup.PrintAreas.Add("A3:B5");
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.DeleteRows(2, 3);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(PrintAreaShouldFollowRowsShiftedUpByDeletionAboveStart), workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Kept1", text, StringComparison.Ordinal);
        Assert.Contains("Kept2", text, StringComparison.Ordinal);
        Assert.DoesNotContain("AbovePrintArea", text, StringComparison.Ordinal);
        Assert.DoesNotContain("BelowPrintArea", text, StringComparison.Ordinal);
    }
}
