// <copyright file="FeatureTests.FontStyle.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>フォントスタイル (太字・斜体) に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldContainBoldText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldText";
            cell.Style.Font.Bold = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainBoldText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainItalicText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ItalicText";
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainItalicText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ItalicText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainBoldItalicText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldItalic";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainBoldItalicText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldItalic", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderMixedNormalBoldItalicOnSamePage()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Normal";
            var boldCell = sheet.Cell("A2");
            boldCell.Value = "Bold";
            boldCell.Style.Font.Bold = true;
            var italicCell = sheet.Cell("A3");
            italicCell.Value = "Italic";
            italicCell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldRenderMixedNormalBoldItalicOnSamePage),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Normal", text, StringComparison.Ordinal);
        Assert.Contains("Bold", text, StringComparison.Ordinal);
        Assert.Contains("Italic", text, StringComparison.Ordinal);
    }
}
