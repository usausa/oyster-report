// <copyright file="FeatureTests.SingleSheetOutput.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>
/// 特定シートのみを対象に PDF を生成する機能テスト。
/// <see cref="OysterReportEngine.GeneratePdf(TemplateSheet, Stream)"/> を使用する。
/// </summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void SingleSheetOutputShouldRenderTargetSheetByIndex()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Sheet1").Cell("A1").Value = "Sheet1Content";
            workbook.AddWorksheet("Sheet2").Cell("A1").Value = "Sheet2Content";
            workbook.AddWorksheet("Sheet3").Cell("A1").Value = "Sheet3Content";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            nameof(SingleSheetOutputShouldRenderTargetSheetByIndex),
            workbook.Sheets[1]);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, PdfTestHelper.GetPageCount(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Sheet2Content", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Sheet1Content", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Sheet3Content", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SingleSheetOutputShouldRenderTargetSheetByName()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Summary").Cell("A1").Value = "SummaryContent";
            workbook.AddWorksheet("Detail").Cell("A1").Value = "DetailContent";
            workbook.AddWorksheet("Appendix").Cell("A1").Value = "AppendixContent";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            nameof(SingleSheetOutputShouldRenderTargetSheetByName),
            workbook.GetSheet("Detail"));

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, PdfTestHelper.GetPageCount(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("DetailContent", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SummaryContent", text, StringComparison.Ordinal);
        Assert.DoesNotContain("AppendixContent", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SingleSheetOutputShouldRenderFirstSheetByIndex0()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("First").Cell("A1").Value = "FirstSheet";
            workbook.AddWorksheet("Second").Cell("A1").Value = "SecondSheet";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            nameof(SingleSheetOutputShouldRenderFirstSheetByIndex0),
            workbook.Sheets[0]);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, PdfTestHelper.GetPageCount(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("FirstSheet", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SecondSheet", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SingleSheetOutputShouldIsolateReplacementsToTargetSheet()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Cover").Cell("A1").Value = "{{Title}}";
            workbook.AddWorksheet("Body").Cell("A1").Value = "{{Content}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var coverSheet = workbook.GetSheet("Cover");
        coverSheet.ReplacePlaceholder("Title", "ReplacedTitle");

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            nameof(SingleSheetOutputShouldIsolateReplacementsToTargetSheet),
            coverSheet);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ReplacedTitle", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
