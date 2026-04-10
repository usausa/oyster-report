// <copyright file="FeatureTests.MergedCell.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>セル結合に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldContainTextInHorizontallyMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "HorizontalMerge";
            sheet.Range("A1:D1").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldContainTextInHorizontallyMergedCell),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("HorizontalMerge", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainTextInVerticallyMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VerticalMerge";
            sheet.Range("A1:A4").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldContainTextInVerticallyMergedCell),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("VerticalMerge", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainTextInRectangularMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("B2").Value = "RectMerge";
            sheet.Range("B2:D4").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldContainTextInRectangularMergedCell),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("RectMerge", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderMultipleMergedRanges()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Range("A1:C1").Merge();
            sheet.Cell("A2").Value = "Left";
            sheet.Range("A2:A4").Merge();
            sheet.Cell("B2").Value = "Right";
            sheet.Range("B2:C4").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldRenderMultipleMergedRanges), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("Left", text, StringComparison.Ordinal);
        Assert.Contains("Right", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldNotDuplicateTextFromMergedSubCells()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "MergeOwner";
            sheet.Range("A1:C1").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldNotDuplicateTextFromMergedSubCells),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var count = CountSubstringOccurrences(PdfTestHelper.ExtractAllText(pdfBytes), "MergeOwner");
        Assert.Equal(1, count);
    }

    private static int CountSubstringOccurrences(string source, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = source.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }
}
