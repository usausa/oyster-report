// <copyright file="ExcelReaderTests.cs" company="machi_pon">
// Copyright (c) machi_pon. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;
using OysterReport.Reading;
using Xunit;

public sealed class ExcelReaderTests
{
    [Fact]
    public void ReadShouldPopulateWorkbookModel()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Title";
            sheet.Cell("A2").Value = "{{Name}}";
            sheet.Range("A1:B1").Merge();
            sheet.Row(2).Hide();
            sheet.PageSetup.CenterHorizontally = true;
            sheet.PageSetup.Header.Left.AddText("Header", XLHFOccurrence.OddPages);
            sheet.PageSetup.Footer.Left.AddText("Footer", XLHFOccurrence.OddPages);
        });

        var reader = new ExcelReader();
        var workbook = reader.Read(stream);

        var sheet = Assert.Single(workbook.Sheets);
        Assert.Equal("Report", sheet.Name);
        Assert.Single(sheet.MergedRanges);
        Assert.Equal("&LHeader", sheet.HeaderFooter.OddHeader);
        Assert.Equal("&LFooter", sheet.HeaderFooter.OddFooter);
        Assert.True(sheet.PageSetup.CenterHorizontally);
        Assert.Contains(sheet.Rows, row => row.Index == 2 && row.IsHidden);
        Assert.Contains(sheet.Cells, cell => cell.Address == "A2" && cell.Placeholder?.MarkerName == "Name");
    }
}
