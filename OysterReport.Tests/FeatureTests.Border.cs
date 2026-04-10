// <copyright file="FeatureTests.Border.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>罫線スタイルに関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Theory]
    [InlineData(XLBorderStyleValues.Thin, "ThinBorder")]
    [InlineData(XLBorderStyleValues.Medium, "MediumBorder")]
    [InlineData(XLBorderStyleValues.Thick, "ThickBorder")]
    [InlineData(XLBorderStyleValues.Double, "DoubleBorder")]
    [InlineData(XLBorderStyleValues.Dashed, "DashedBorder")]
    [InlineData(XLBorderStyleValues.Dotted, "DottedBorder")]
    public void PdfShouldContainCellWithBorder(XLBorderStyleValues borderStyle, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = cellValue;
            cell.Style.Border.OutsideBorder = borderStyle;
            cell.Style.Border.OutsideBorderColor = XLColor.Black;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(PdfShouldContainCellWithBorder)}_{borderStyle}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderColoredBorder()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = "ColoredBorder";
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            cell.Style.Border.OutsideBorderColor = XLColor.FromArgb(255, 0, 0);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldRenderColoredBorder), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ColoredBorder", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderTableWithAllSideBorders()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            for (var row = 1; row <= 3; row++)
            {
                for (var col = 1; col <= 3; col++)
                {
                    var cell = sheet.Cell(row, col);
                    cell.Value = $"R{row}C{col}";
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldRenderTableWithAllSideBorders), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("R1C1", text, StringComparison.Ordinal);
        Assert.Contains("R3C3", text, StringComparison.Ordinal);
    }
}
