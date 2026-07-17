namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void DetailRowExpansionShouldContainExactlyThreeRows()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{No}}: {{ItemName}}";
            sheet.Cell("A3").Value = "Footer";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("ItemName");

        var last = template;
        foreach (var (no, name) in new[] { ("1", "Apple"), ("2", "Banana"), ("3", "Cherry") })
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholder("No", no);
            last.ReplacePlaceholder("ItemName", name);
        }

        template.Delete();

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldContainExactlyThreeRows),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Apple", text, StringComparison.Ordinal);
        Assert.Contains("Banana", text, StringComparison.Ordinal);
        Assert.Contains("Cherry", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{ItemName}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{No}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Durian", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DetailRowExpansionShouldRemovePlaceholderAfterDeletion()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Product}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Product");
        var row = template.InsertCopyBelow();
        row.ReplacePlaceholder("Product", "Widget");
        template.Delete();

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldRemovePlaceholderAfterDeletion),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Product}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DetailRowExpansionShouldPreserveHeaderAndFooter()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "ReportHeader";
            sheet.Cell("A2").Value = "{{Line}}";
            sheet.Cell("A3").Value = "ReportFooter";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Line");

        var last = template;
        last = template.InsertCopyAfter(last);
        last.ReplacePlaceholder("Line", "LineA");
        last = template.InsertCopyAfter(last);
        last.ReplacePlaceholder("Line", "LineB");
        template.Delete();

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldPreserveHeaderAndFooter),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("ReportHeader", text, StringComparison.Ordinal);
        Assert.Contains("LineA", text, StringComparison.Ordinal);
        Assert.Contains("LineB", text, StringComparison.Ordinal);
        Assert.Contains("ReportFooter", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DetailRowExpansionShouldCountExactOccurrencesOfLabel()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "START";
            sheet.Cell("A2").Value = "ROW-{{Seq}}";
            sheet.Cell("A3").Value = "END";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Seq");

        var last = template;
        for (var i = 1; i <= 4; i++)
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholder("Seq", $"{i}");
        }

        template.Delete();

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldCountExactOccurrencesOfLabel),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("ROW-1", text, StringComparison.Ordinal);
        Assert.Contains("ROW-2", text, StringComparison.Ordinal);
        Assert.Contains("ROW-3", text, StringComparison.Ordinal);
        Assert.Contains("ROW-4", text, StringComparison.Ordinal);
        Assert.DoesNotContain("ROW-5", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Seq}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DetailRowExpansionShouldShareImmutableStyleYetIsolateSourceOnMutation()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "{{Item}}";
            cell.Style.Font.Bold = true;
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var underlying = sheet.UnderlyingSheet;
        var template = sheet.FindRow("Item");
        var sourceCell = underlying.FindCell(1, 1);
        Assert.NotNull(sourceCell);

        // Act — copy the row, then replace the placeholder only in the copy
        var copy = template.InsertCopyBelow();
        copy.ReplacePlaceholder("Item", "Widget");
        var copiedCell = underlying.FindCell(2, 1);

        // Assert
        Assert.NotNull(copiedCell);
        // The immutable style is shared with the source (no per-copy allocation)
        Assert.Same(sourceCell.Style, copiedCell.Style);
        // Replacing the copy's text swaps its whole value and leaves the source untouched
        Assert.Equal("{{Item}}", sourceCell.DisplayText);
        Assert.Equal("Widget", copiedCell.DisplayText);
        Assert.NotSame(sourceCell.Value, copiedCell.Value);
    }
}
