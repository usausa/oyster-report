// <copyright file="ReportSheetTests.cs" company="machi_pon">
// Copyright (c) machi_pon. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Model;
using OysterReport.Reading;
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

        var reader = new ExcelReader();
        var workbook = reader.Read(stream);
        var sheet = Assert.Single(workbook.Sheets);

        var count = sheet.ReplacePlaceholder("CustomerName", "Alice");

        Assert.Equal(1, count);
        Assert.Equal("Alice", sheet.Cells.Single(cell => cell.Address == "A1").DisplayText);
    }

    [Fact]
    public void AddRowsShouldDuplicateTemplateRowsAndShiftFollowingRows()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
            sheet.Cell("A3").Value = "Footer";
        });

        var reader = new ExcelReader();
        var workbook = reader.Read(stream);
        var sheet = Assert.Single(workbook.Sheets);

        sheet.AddRows(new RowExpansionRequest
        {
            TemplateStartRowIndex = 2,
            TemplateEndRowIndex = 2,
            RepeatCount = 2,
            PlaceholderValuesByIteration =
            [
                new Dictionary<string, string?> { ["Item"] = "A" },
                new Dictionary<string, string?> { ["Item"] = "B" },
            ],
        });

        Assert.Contains(sheet.Cells, cell => cell.Address == "A3" && cell.DisplayText == "A");
        Assert.Contains(sheet.Cells, cell => cell.Address == "A4" && cell.DisplayText == "B");
        Assert.Contains(sheet.Cells, cell => cell.Address == "A5" && cell.DisplayText == "Footer");
        Assert.Equal(5, sheet.UsedRange.EndRow);
    }
}
