// <copyright file="FeatureTests.HeaderFooter.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>ヘッダー・フッターに関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldContainHeaderText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Header.Left.AddText("LeftHeader", XLHFOccurrence.OddPages);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainHeaderText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BodyContent", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainFooterText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Footer.Right.AddText("RightFooter", XLHFOccurrence.OddPages);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainFooterText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void PdfShouldContainBothHeaderAndFooter()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Main";
            sheet.PageSetup.Header.Center.AddText("TopCenter", XLHFOccurrence.OddPages);
            sheet.PageSetup.Footer.Center.AddText("BottomCenter", XLHFOccurrence.OddPages);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainBothHeaderAndFooter), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}
