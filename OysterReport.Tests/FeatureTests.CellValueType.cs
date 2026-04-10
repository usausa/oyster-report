// <copyright file="FeatureTests.CellValueType.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>セルの値型 (文字列・数値・日付・数式) に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldRenderNumericCellValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = 12345;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldRenderNumericCellValue), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("12345", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderDateCellValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = new DateTime(2025, 1, 15);
            cell.Style.DateFormat.Format = "yyyy/MM/dd";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldRenderDateCellValue), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("2025", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderFormulaCellValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = 10;
            sheet.Cell("A2").Value = 20;
            sheet.Cell("A3").FormulaA1 = "=A1+A2";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldRenderFormulaCellValue), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}
