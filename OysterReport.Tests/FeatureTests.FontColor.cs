// <copyright file="FeatureTests.FontColor.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>フォントカラーに関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Theory]
    [InlineData("RedText", 255, 0, 0)]
    [InlineData("BlueText", 0, 0, 255)]
    [InlineData("GreenText", 0, 128, 0)]
    public void PdfShouldContainColoredText(string cellValue, int r, int g, int b)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Font.FontColor = XLColor.FromArgb(r, g, b);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(PdfShouldContainColoredText)}_{cellValue}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainThemeColorText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeColorText";
            cell.Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.4);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainThemeColorText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeColorText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
