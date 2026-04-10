// <copyright file="FeatureTests.HiddenRowColumn.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>非表示行・列に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldNotContainHiddenRowText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VisibleRow";
            sheet.Cell("A2").Value = "HiddenRow";
            sheet.Row(2).Hide();
            sheet.Cell("A3").Value = "VisibleRow2";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldNotContainHiddenRowText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("VisibleRow", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HiddenRow", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldNotContainHiddenColumnText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VisibleCol";
            sheet.Cell("B1").Value = "HiddenCol";
            sheet.Column(2).Hide();
            sheet.Cell("C1").Value = "VisibleCol2";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldNotContainHiddenColumnText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("VisibleCol", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HiddenCol", text, StringComparison.Ordinal);
    }
}
