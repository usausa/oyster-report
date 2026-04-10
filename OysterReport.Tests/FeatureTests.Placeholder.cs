// <copyright file="FeatureTests.Placeholder.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>プレースホルダー置換に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldContainReplacedPlaceholderValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{CustomerName}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).ReplacePlaceholder("CustomerName", "AcmeCorp");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldContainReplacedPlaceholderValue),
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("AcmeCorp", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainAllReplacedPlaceholders()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Title}}";
            sheet.Cell("A2").Value = "{{Name}}";
            sheet.Cell("A3").Value = "{{Date}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["Title"] = "Invoice",
            ["Name"] = "JohnDoe",
            ["Date"] = "2025-01-01"
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldContainAllReplacedPlaceholders),
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Invoice", text, StringComparison.Ordinal);
        Assert.Contains("JohnDoe", text, StringComparison.Ordinal);
        Assert.Contains("2025-01-01", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldPreservePlaceholderNotReplaced()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{KeepMe}}";
            sheet.Cell("A2").Value = "{{ReplaceMe}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).ReplacePlaceholder("ReplaceMe", "Replaced");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldPreservePlaceholderNotReplaced),
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("Replaced", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainReplacedPlaceholderInMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{MergedTitle}}";
            sheet.Range("A1:C1").Merge();
        });

        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).ReplacePlaceholder("MergedTitle", "MergedValue");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldContainReplacedPlaceholderInMergedCell),
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("MergedValue", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
