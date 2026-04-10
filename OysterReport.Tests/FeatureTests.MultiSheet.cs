// <copyright file="FeatureTests.MultiSheet.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>複数シートに関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldHaveOnePagePerSheet()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Sheet1").Cell("A1").Value = "ContentSheet1";
            workbook.AddWorksheet("Sheet2").Cell("A1").Value = "ContentSheet2";
            workbook.AddWorksheet("Sheet3").Cell("A1").Value = "ContentSheet3";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldHaveOnePagePerSheet), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.True(PdfTestHelper.GetPageCount(pdfBytes) >= 3);
    }

    [Fact]
    public void PdfShouldContainTextFromAllSheets()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Alpha").Cell("A1").Value = "AlphaSheet";
            workbook.AddWorksheet("Beta").Cell("A1").Value = "BetaSheet";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainTextFromAllSheets), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("AlphaSheet", text, StringComparison.Ordinal);
        Assert.Contains("BetaSheet", text, StringComparison.Ordinal);
    }
}
