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
    public void FontColorShouldRenderColoredText(string cellValue, int r, int g, int b)
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Font.FontColor = XLColor.FromArgb(r, g, b);
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(FontColorShouldRenderColoredText)}_{cellValue}",
            stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontColorShouldRenderThemeColorText()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeColorText";
            cell.Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.4);
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontColorShouldRenderThemeColorText), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeColorText", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
