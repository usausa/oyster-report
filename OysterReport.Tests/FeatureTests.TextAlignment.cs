// <copyright file="FeatureTests.TextAlignment.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>テキスト配置 (水平・垂直・折り返し) に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Theory]
    [InlineData(XLAlignmentHorizontalValues.Left, "LeftAligned")]
    [InlineData(XLAlignmentHorizontalValues.Center, "CenterAligned")]
    [InlineData(XLAlignmentHorizontalValues.Right, "RightAligned")]
    public void PdfShouldContainHorizontallyAlignedText(XLAlignmentHorizontalValues alignment, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Alignment.Horizontal = alignment;
            sheet.Column(1).Width = 30d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(PdfShouldContainHorizontallyAlignedText)}_{alignment}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(XLAlignmentVerticalValues.Top, "TopAligned")]
    [InlineData(XLAlignmentVerticalValues.Center, "MiddleAligned")]
    [InlineData(XLAlignmentVerticalValues.Bottom, "BottomAligned")]
    public void PdfShouldContainVerticallyAlignedText(XLAlignmentVerticalValues alignment, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Alignment.Vertical = alignment;
            sheet.Row(1).Height = 40d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(PdfShouldContainVerticallyAlignedText)}_{alignment}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainWrappedText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "This is a long text that should wrap inside the cell boundaries";
            cell.Style.Alignment.WrapText = true;
            sheet.Column(1).Width = 20d;
            sheet.Row(1).Height = 50d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldContainWrappedText), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}
