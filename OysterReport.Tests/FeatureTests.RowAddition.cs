namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void RowAdditionShouldContainRowsFromInsertCopyBelow()
    {
        // Arrange
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

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromInsertCopyBelow),
            workbook);

        // Assert
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
        // Arrange
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

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromInsertCopyAfter),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Row1", text, StringComparison.Ordinal);
        Assert.Contains("Row2", text, StringComparison.Ordinal);
        Assert.Contains("Row3", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldPreserveStyleAfterCopy()
    {
        // Arrange
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

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(RowAdditionShouldPreserveStyleAfterCopy), workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("CopiedRow", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldHandleZeroDetailRows()
    {
        // Arrange
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

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(RowAdditionShouldHandleZeroDetailRows), workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("Footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldContainRowsFromMultiRowRangeExpansion()
    {
        // Arrange
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

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromMultiRowRangeExpansion),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alice", text, StringComparison.Ordinal);
        Assert.Contains("Detail1", text, StringComparison.Ordinal);
        Assert.Contains("Bob", text, StringComparison.Ordinal);
        Assert.Contains("Detail2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldContainRowsFromInsertCopyBelowOnLastRow()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
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
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromInsertCopyBelowOnLastRow),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("ItemA", text, StringComparison.Ordinal);
        Assert.Contains("ItemB", text, StringComparison.Ordinal);
        Assert.Contains("ItemC", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldContainRowsFromMultiRowRangeExpansionOnLastRows()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Name}}";
            sheet.Cell("A3").Value = "{{Detail}}";
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

        templateRange.Delete();

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsFromMultiRowRangeExpansionOnLastRows),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alice", text, StringComparison.Ordinal);
        Assert.Contains("Detail1", text, StringComparison.Ordinal);
        Assert.Contains("Bob", text, StringComparison.Ordinal);
        Assert.Contains("Detail2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowAdditionShouldContainRowsShiftedUpByDeletionAboveUsedRange()
    {
        // Arrange — the used range starts at row 2 and the deletion starts above it
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A2").Value = "Deleted";
            sheet.Cell("A3").Value = "Kept1";
            sheet.Cell("A4").Value = "Kept2";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.DeleteRows(1, 2);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(RowAdditionShouldContainRowsShiftedUpByDeletionAboveUsedRange),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Kept1", text, StringComparison.Ordinal);
        Assert.Contains("Kept2", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", text, StringComparison.Ordinal);
    }
}
