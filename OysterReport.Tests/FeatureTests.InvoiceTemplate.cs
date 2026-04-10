// <copyright file="FeatureTests.InvoiceTemplate.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>
/// 請求書形式の総合テンプレートに関する機能テスト。
/// 複数機能を組み合わせた実帳票シナリオを確認する。
/// </summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void PdfShouldRenderInvoiceTemplateWithAllFeatures()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Invoice");
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet.PageSetup.Margins.Left = 1.8;
            sheet.PageSetup.Margins.Right = 1.8;
            sheet.PageSetup.Margins.Top = 2.5;
            sheet.PageSetup.Margins.Bottom = 2.5;

            sheet.Cell("A1").Value = "{{Title}}";
            sheet.Range("A1:F1").Merge();
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 16d;
            sheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            sheet.Cell("A3").Value = "宛先:";
            sheet.Cell("B3").Value = "{{CustomerName}}";
            sheet.Cell("A4").Value = "日付:";
            sheet.Cell("B4").Value = "{{IssueDate}}";

            foreach (var (col, label) in new[] { ("A6", "品目"), ("B6", "数量"), ("C6", "単価"), ("D6", "金額") })
            {
                sheet.Cell(col).Value = label;
                sheet.Cell(col).Style.Font.Bold = true;
                sheet.Cell(col).Style.Fill.BackgroundColor = XLColor.FromArgb(220, 220, 220);
                sheet.Cell(col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            foreach (var (col, placeholder) in new[] { ("A7", "ItemName"), ("B7", "Qty"), ("C7", "UnitPrice"), ("D7", "Amount") })
            {
                sheet.Cell(col).Value = $"{{{{{placeholder}}}}}";
                sheet.Cell(col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            sheet.Cell("A9").Value = "{{FooterNote}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        sheet.ReplacePlaceholder("Title", "請求書");
        sheet.ReplacePlaceholder("CustomerName", "株式会社サンプル");
        sheet.ReplacePlaceholder("IssueDate", "2025-01-15");

        var templateRow = sheet.FindRow("ItemName");
        var lastRow = templateRow;
        foreach (var (name, qty, price, amount) in new[]
        {
            ("商品A", "2", "1000", "2000"),
            ("商品B", "1", "3000", "3000"),
            ("商品C", "5", "500", "2500")
        })
        {
            lastRow = templateRow.InsertCopyAfter(lastRow);
            lastRow.ReplacePlaceholder("ItemName", name);
            lastRow.ReplacePlaceholder("Qty", qty);
            lastRow.ReplacePlaceholder("UnitPrice", price);
            lastRow.ReplacePlaceholder("Amount", amount);
        }

        templateRow.Delete();
        sheet.ReplacePlaceholder("FooterNote", "上記の通りご請求申し上げます。");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(PdfShouldRenderInvoiceTemplateWithAllFeatures),
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("2025-01-15", text, StringComparison.Ordinal);
        Assert.Contains("2000", text, StringComparison.Ordinal);
        Assert.Contains("3000", text, StringComparison.Ordinal);
        Assert.Contains("2500", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderListReportWithStripedRows()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("List");
            sheet.Cell("A1").Value = "No";
            sheet.Cell("B1").Value = "Name";
            sheet.Cell("C1").Value = "Score";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("B1").Style.Font.Bold = true;
            sheet.Cell("C1").Style.Font.Bold = true;

            for (var i = 1; i <= 10; i++)
            {
                var bgColor = i % 2 == 0
                    ? XLColor.FromArgb(240, 248, 255)
                    : XLColor.FromArgb(255, 255, 255);
                sheet.Cell(i + 1, 1).Value = i;
                sheet.Cell(i + 1, 2).Value = $"Name{i:D2}";
                sheet.Cell(i + 1, 3).Value = i * 10;
                for (var col = 1; col <= 3; col++)
                {
                    sheet.Cell(i + 1, col).Style.Fill.BackgroundColor = bgColor;
                    sheet.Cell(i + 1, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(nameof(PdfShouldRenderListReportWithStripedRows), stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Name01", text, StringComparison.Ordinal);
        Assert.Contains("Name10", text, StringComparison.Ordinal);
    }
}
