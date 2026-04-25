namespace OysterReport.Tests;

public sealed class TemplateWorkbookCloneTests
{
    private static readonly byte[] OnePxPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");

    [Fact]
    public void CloneShouldReturnSeparateTemplateAndReportWorkbookInstances()
    {
        // Arrange
        using var template = CreateRichTemplate();

        // Act
        using var clone = template.Clone();

        // Assert
        Assert.NotSame(template, clone);
        Assert.NotSame(template.ReportWorkbook, clone.ReportWorkbook);
        Assert.NotSame(template.Sheets, clone.Sheets);
        Assert.NotSame(template.ReportWorkbook.Metadata, clone.ReportWorkbook.Metadata);
        Assert.NotSame(template.ReportWorkbook.MeasurementProfile, clone.ReportWorkbook.MeasurementProfile);
        Assert.Equal(template.ReportWorkbook.Metadata, clone.ReportWorkbook.Metadata);
        Assert.Equal(template.ReportWorkbook.MeasurementProfile, clone.ReportWorkbook.MeasurementProfile);
        Assert.Equal(template.Sheets.Count, clone.Sheets.Count);
    }

    [Fact]
    public void CloneShouldCreateSeparateSheetInstancesAtEveryLevel()
    {
        // Arrange
        using var template = CreateRichTemplate();

        // Act
        using var clone = template.Clone();

        // Assert
        Assert.Equal(template.Sheets.Count, clone.Sheets.Count);
        for (var i = 0; i < template.Sheets.Count; i++)
        {
            var a = template.Sheets[i];
            var b = clone.Sheets[i];

            Assert.NotSame(a, b);
            Assert.NotSame(a.UnderlyingSheet, b.UnderlyingSheet);
            Assert.Equal(a.Name, b.Name);

            var sheetA = a.UnderlyingSheet;
            var sheetB = b.UnderlyingSheet;

            Assert.NotSame(sheetA.PageSetup, sheetB.PageSetup);
            Assert.Equal(sheetA.PageSetup, sheetB.PageSetup);

            Assert.NotSame(sheetA.HeaderFooter, sheetB.HeaderFooter);
            Assert.Equal(sheetA.HeaderFooter, sheetB.HeaderFooter);

            if (sheetA.PrintArea is not null)
            {
                Assert.NotNull(sheetB.PrintArea);
                Assert.NotSame(sheetA.PrintArea, sheetB.PrintArea);
                Assert.Equal(sheetA.PrintArea, sheetB.PrintArea);
            }

            AssertSeparateCollectionInstances(sheetA.Rows, sheetB.Rows);
            AssertSeparateCollectionInstances(sheetA.Columns, sheetB.Columns);
            AssertSeparateCollectionInstances(sheetA.MergedRanges, sheetB.MergedRanges);
            AssertSeparateCollectionInstances(sheetA.Images, sheetB.Images);
            AssertSeparateCollectionInstances(sheetA.HorizontalPageBreaks, sheetB.HorizontalPageBreaks);
            AssertSeparateCollectionInstances(sheetA.VerticalPageBreaks, sheetB.VerticalPageBreaks);

            Assert.Equal(sheetA.Cells.Count, sheetB.Cells.Count);
            for (var c = 0; c < sheetA.Cells.Count; c++)
            {
                var cellA = sheetA.Cells[c];
                var cellB = sheetB.Cells[c];

                Assert.NotSame(cellA, cellB);
                Assert.Equal(cellA.Row, cellB.Row);
                Assert.Equal(cellA.Column, cellB.Column);
                Assert.Equal(cellA.DisplayText, cellB.DisplayText);

                Assert.NotSame(cellA.Value, cellB.Value);
                Assert.Equal(cellA.Value, cellB.Value);

                Assert.NotSame(cellA.Style, cellB.Style);
                Assert.Equal(cellA.Style, cellB.Style);

                Assert.NotSame(cellA.Style.Font, cellB.Style.Font);
                Assert.NotSame(cellA.Style.Fill, cellB.Style.Fill);
                Assert.NotSame(cellA.Style.Alignment, cellB.Style.Alignment);

                Assert.NotSame(cellA.Style.Borders, cellB.Style.Borders);
                Assert.NotSame(cellA.Style.Borders.Left, cellB.Style.Borders.Left);
                Assert.NotSame(cellA.Style.Borders.Top, cellB.Style.Borders.Top);
                Assert.NotSame(cellA.Style.Borders.Right, cellB.Style.Borders.Right);
                Assert.NotSame(cellA.Style.Borders.Bottom, cellB.Style.Borders.Bottom);

                if (cellA.Merge is not null)
                {
                    Assert.NotNull(cellB.Merge);
                    Assert.NotSame(cellA.Merge, cellB.Merge);
                    Assert.Equal(cellA.Merge, cellB.Merge);
                }
            }
        }
    }

    [Fact]
    public void CloneShouldIsolateMutationsFromOriginal()
    {
        // Arrange
        using var template = CreateRichTemplate();
        var originalSheet = template.Sheets[0];
        var originalCell = originalSheet.UnderlyingSheet.Cells[0];
        var originalText = originalCell.DisplayText;
        var originalCellValue = originalCell.Value;
        var originalCellStyle = originalCell.Style;

        // Act
        using var clone = template.Clone();
        var clonedSheet = clone.Sheets[0];
        var clonedCell = clonedSheet.UnderlyingSheet.Cells[0];
        TemplateSheet.SetCellText(clonedCell, "Mutated");
        clonedSheet.UnderlyingSheet.DeleteRows(2, 2);

        // Assert
        Assert.Equal(originalText, originalCell.DisplayText);
        Assert.Same(originalCellValue, originalCell.Value);
        Assert.Same(originalCellStyle, originalCell.Style);
        Assert.NotEqual(originalSheet.UnderlyingSheet.Cells.Count, clonedSheet.UnderlyingSheet.Cells.Count);
    }

    private static TemplateWorkbook CreateRichTemplate()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet1 = workbook.AddWorksheet("Report");
            sheet1.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet1.PageSetup.Margins.Left = 0.6;
            sheet1.PageSetup.Header.Center.AddText("HDR");
            sheet1.PageSetup.Footer.Center.AddText("FTR");
            sheet1.PageSetup.PrintAreas.Add("A1:C3");

            sheet1.Column(1).Width = 20;
            sheet1.Column(2).Width = 10;
            sheet1.Row(1).Height = 30;

            var a1 = sheet1.Cell("A1");
            a1.Value = "Title";
            a1.Style.Font.FontName = "Meiryo";
            a1.Style.Font.FontSize = 14;
            a1.Style.Font.Bold = true;
            a1.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 240, 240, 240);
            a1.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            a1.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            a1.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            a1.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            a1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            a1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            sheet1.Cell("A2").Value = "Body";
            sheet1.Cell("B2").Value = 123;
            sheet1.Cell("C2").Value = "{{Placeholder}}";

            sheet1.Range("A3:B3").Merge();
            sheet1.Cell("A3").Value = "Merged";

            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet1.AddPicture(imgStream, XLPictureFormat.Png, "Pic").MoveTo(sheet1.Cell("D1")).WithSize(40, 30);

            var sheet2 = workbook.AddWorksheet("Second");
            sheet2.Cell("A1").Value = "OtherSheet";
        });

        return new TemplateWorkbook(stream);
    }

    private static void AssertSeparateCollectionInstances<T>(IReadOnlyList<T> source, IReadOnlyList<T> clone)
        where T : class
    {
        Assert.NotSame(source, clone);
        Assert.Equal(source.Count, clone.Count);
        for (var i = 0; i < source.Count; i++)
        {
            Assert.NotSame(source[i], clone[i]);
        }
    }
}
