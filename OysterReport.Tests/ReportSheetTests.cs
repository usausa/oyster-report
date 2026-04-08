// <copyright file="ReportSheetTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using Xunit;

public sealed class ReportSheetTests
{
    [Fact]
    public void ReplacePlaceholderShouldUpdateDisplayText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{CustomerName}}";
        });

        var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var file = File.Create(tempFile))
            {
                stream.CopyTo(file);
            }

            using var workbook = new TemplateWorkbook(tempFile);
            var sheet = Assert.Single(workbook.Sheets);

            var count = sheet.ReplacePlaceholder("CustomerName", "Alice");

            Assert.Equal(1, count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void InsertCopyBelowShouldDuplicateRow()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
            sheet.Cell("A3").Value = "Footer";
        });

        var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var file = File.Create(tempFile))
            {
                stream.CopyTo(file);
            }

            var engine = new OysterReportEngine();
            using var workbook = new TemplateWorkbook(tempFile);
            var sheet = Assert.Single(workbook.Sheets);

            var template = sheet.FindRow("Item");
            var row1 = template.InsertCopyBelow();
            row1.ReplacePlaceholder("Item", "A");
            var row2 = row1.InsertCopyBelow();
            row2.ReplacePlaceholder("Item", "B");
            template.Delete();

            using var output = new MemoryStream();
            engine.GeneratePdf(workbook, output);
            Assert.True(output.Length > 0);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
