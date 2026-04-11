// <copyright file="FeatureTests.FontSize.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>フォントサイズに関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Theory]
    [InlineData(6d, "TinyText")]
    [InlineData(11d, "NormalText")]
    [InlineData(18d, "LargeText")]
    [InlineData(24d, "HugeText")]
    public void FontSizeShouldRenderVariousSizes(double fontSize, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Font.FontSize = fontSize;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FontSizeShouldRenderVariousSizes)}_{fontSize}pt",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontSizeShouldRenderMultipleSizesOnOnePage()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Size8";
            sheet.Cell("A1").Style.Font.FontSize = 8d;
            sheet.Cell("A2").Value = "Size12";
            sheet.Cell("A2").Style.Font.FontSize = 12d;
            sheet.Cell("A3").Value = "Size16";
            sheet.Cell("A3").Style.Font.FontSize = 16d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(FontSizeShouldRenderMultipleSizesOnOnePage),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Size8", text, StringComparison.Ordinal);
        Assert.Contains("Size12", text, StringComparison.Ordinal);
        Assert.Contains("Size16", text, StringComparison.Ordinal);
    }
}
