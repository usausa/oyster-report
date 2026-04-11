namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>一括プレースホルダー置換に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void ReplacePlaceholdersShouldReplaceMultipleOnRow()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "name:{{PersonName}} dept:{{PersonDept}} city:{{PersonCity}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.GetRow(1).ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["PersonName"] = "tanaka",
            ["PersonDept"] = "sales",
            ["PersonCity"] = "tokyo"
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldReplaceMultipleOnRow),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("tanaka", text, StringComparison.Ordinal);
        Assert.Contains("sales", text, StringComparison.Ordinal);
        Assert.Contains("tokyo", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersShouldTreatNullValueAsEmpty()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Name: {{Name}}";
            sheet.Cell("B1").Value = "Memo: {{Memo}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).GetRow(1).ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["Name"] = "Alice",
            ["Memo"] = null
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldTreatNullValueAsEmpty),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alice", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Memo}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersShouldReplaceMultipleOnRowRange()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Item: {{ItemName}}";
            sheet.Cell("A2").Value = "Price: {{Price}}";
            sheet.Cell("A3").Value = "Qty: {{Qty}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).FindRows("ItemName").ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["ItemName"] = "Widget",
            ["Price"] = "980",
            ["Qty"] = "5"
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldReplaceMultipleOnRowRange),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("980", text, StringComparison.Ordinal);
        Assert.Contains("5", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersShouldReplaceAcrossAllSheets()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Cover").Cell("A1").Value = "{{DocTitle}}";
            workbook.AddWorksheet("Body").Cell("A1").Value = "Author: {{Author}}";
            workbook.AddWorksheet("Appendix").Cell("A1").Value = "{{DocTitle}} - Appendix";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        workbook.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["DocTitle"] = "AnnualReport",
            ["Author"] = "Smith"
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldReplaceAcrossAllSheets),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("AnnualReport", text, StringComparison.Ordinal);
        Assert.Contains("Smith", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{DocTitle}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Author}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersShouldWorkWithExpandedRowsInLoop()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "HEADER";
            sheet.Cell("A2").Value = "{{Code}}";
            sheet.Cell("B2").Value = "{{Label}}";
            sheet.Cell("C2").Value = "{{Value}}";
            sheet.Cell("A3").Value = "FOOTER";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Code");

        var items = new[]
        {
            new Dictionary<string, string?> { ["Code"] = "001", ["Label"] = "Alpha", ["Value"] = "100" },
            new Dictionary<string, string?> { ["Code"] = "002", ["Label"] = "Beta",  ["Value"] = "200" },
            new Dictionary<string, string?> { ["Code"] = "003", ["Label"] = "Gamma", ["Value"] = "300" }
        };

        var last = template;
        foreach (var item in items)
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholders(item);
        }

        template.Delete();

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldWorkWithExpandedRowsInLoop),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("001", text, StringComparison.Ordinal);
        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("002", text, StringComparison.Ordinal);
        Assert.Contains("Beta", text, StringComparison.Ordinal);
        Assert.Contains("003", text, StringComparison.Ordinal);
        Assert.Contains("Gamma", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Code}}", text, StringComparison.Ordinal);
    }
}
