namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Fact]
    public void DetailRowExpansionShouldContainExactlyThreeRows()
    {
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

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldContainExactlyThreeRows),
            workbook);

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

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldRemovePlaceholderAfterDeletion),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Product}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DetailRowExpansionShouldPreserveHeaderAndFooter()
    {
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

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldPreserveHeaderAndFooter),
            workbook);

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

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(DetailRowExpansionShouldCountExactOccurrencesOfLabel),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("ROW-1", text, StringComparison.Ordinal);
        Assert.Contains("ROW-2", text, StringComparison.Ordinal);
        Assert.Contains("ROW-3", text, StringComparison.Ordinal);
        Assert.Contains("ROW-4", text, StringComparison.Ordinal);
        Assert.DoesNotContain("ROW-5", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Seq}}", text, StringComparison.Ordinal);
    }
}
