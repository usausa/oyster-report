// <copyright file="FeatureTests.PageSetup.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>ページサイズ・余白・中央配置などのページ設定に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PageSetupShouldProduceA4Dimensions()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "A4Page";
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(PageSetupShouldProduceA4Dimensions), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var (width, height) = TestHelper.GetPageSize(pdfBytes);
        Assert.Equal(595.28d, width, 0);
        Assert.Equal(841.89d, height, 0);
    }

    [Fact]
    public void PageSetupShouldProduceA4LandscapeDimensions()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "A4Landscape";
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(PageSetupShouldProduceA4LandscapeDimensions), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var (width, height) = TestHelper.GetPageSize(pdfBytes);
        Assert.Equal(841.89d, width, 0);
        Assert.Equal(595.28d, height, 0);
    }

    [Fact]
    public void PageSetupShouldCenterHorizontally()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Centered";
            sheet.PageSetup.CenterHorizontally = true;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(PageSetupShouldCenterHorizontally), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("Centered", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PageSetupShouldGenerateMultiplePagesForOverflow()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            for (var row = 1; row <= 60; row++)
            {
                sheet.Cell(row, 1).Value = $"Row{row:D2}";
                sheet.Row(row).Height = 20d;
            }
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(PageSetupShouldGenerateMultiplePagesForOverflow),
            stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, TestHelper.GetPageCount(pdfBytes));
        Assert.Contains("Row01", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
        Assert.Contains("Row60", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PageSetupShouldApplyManualPageBreak()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Page1Content";
            sheet.Cell("A2").Value = "Page2Content";
            sheet.PageSetup.AddHorizontalPageBreak(1);
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(PageSetupShouldApplyManualPageBreak), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Page1Content", text, StringComparison.Ordinal);
        Assert.Contains("Page2Content", text, StringComparison.Ordinal);
    }
}
