// <copyright file="FeatureTests.PrintArea.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>印刷範囲に関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldOnlyContainTextWithinPrintArea()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "InPrintArea";
            sheet.Cell("A2").Value = "InPrintArea2";
            sheet.Cell("A10").Value = "OutsidePrintArea";
            sheet.PageSetup.PrintAreas.Add("A1:B5");
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldOnlyContainTextWithinPrintArea), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("InPrintArea", text, StringComparison.Ordinal);
        Assert.DoesNotContain("OutsidePrintArea", text, StringComparison.Ordinal);
    }
}
