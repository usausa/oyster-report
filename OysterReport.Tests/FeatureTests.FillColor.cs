// <copyright file="FeatureTests.FillColor.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>セル背景色に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Theory]
    [InlineData("YellowBg", 255, 255, 0)]
    [InlineData("LightBlueBg", 173, 216, 230)]
    [InlineData("GrayBg", 192, 192, 192)]
    public void FillColorShouldRenderTextOnColoredBackground(string cellValue, int r, int g, int b)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(r, g, b);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FillColorShouldRenderTextOnColoredBackground)}_{cellValue}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FillColorShouldRenderMultipleDifferentBackgroundColors()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var redCell = sheet.Cell("A1");
            redCell.Value = "RedBg";
            redCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 200, 200);
            var greenCell = sheet.Cell("A2");
            greenCell.Value = "GreenBg";
            greenCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 255, 200);
            var blueCell = sheet.Cell("A3");
            blueCell.Value = "BlueBg";
            blueCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 200, 255);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(FillColorShouldRenderMultipleDifferentBackgroundColors),
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("RedBg", text, StringComparison.Ordinal);
        Assert.Contains("GreenBg", text, StringComparison.Ordinal);
        Assert.Contains("BlueBg", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FillColorShouldRenderThemeBackgroundColor()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeBg";
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.25);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(FillColorShouldRenderThemeBackgroundColor), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeBg", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
