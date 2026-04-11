// <copyright file="FeatureTests.FontStyle.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>フォントスタイル (太字・斜体・下線・打ち消し線) に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void FontStyleShouldRenderBoldText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldText";
            cell.Style.Font.Bold = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderBoldText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderItalicText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ItalicText";
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderItalicText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ItalicText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderBoldItalicText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldItalic";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderBoldItalicText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldItalic", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderMixedStylesOnSamePage()
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
            nameof(FontStyleShouldRenderMixedStylesOnSamePage),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Normal", text, StringComparison.Ordinal);
        Assert.Contains("Bold", text, StringComparison.Ordinal);
        Assert.Contains("Italic", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderUnderlinedText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "UnderlinedText";
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderUnderlinedText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("UnderlinedText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderStrikethroughText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "StrikethroughText";
            cell.Style.Font.Strikethrough = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderStrikethroughText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("StrikethroughText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderAllDecorationsOnSameCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "AllDecorations";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Italic = true;
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
            cell.Style.Font.Strikethrough = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderAllDecorationsOnSameCell), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("AllDecorations", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
