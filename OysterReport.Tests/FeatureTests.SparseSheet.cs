namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void SparseSheetShouldNotMaterializeBlankCellsAcrossWidePrintArea()
    {
        // Arrange — real data in two cells but a print area spanning 10,000 empty coordinates
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("B2").Value = "Body";
            sheet.PageSetup.PrintAreas.Add("A1:J1000");
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Assert — only the two real cells are materialized, not the full A1:J1000 grid
        Assert.Equal(2, sheet.UnderlyingSheet.Cells.Count);
        Assert.NotNull(sheet.UnderlyingSheet.FindCell(1, 1));
        Assert.NotNull(sheet.UnderlyingSheet.FindCell(2, 2));
        Assert.Null(sheet.UnderlyingSheet.FindCell(5, 5));

        // Rows are sparse too: only the rows carrying the two cells are materialized
        Assert.Equal(2, sheet.UnderlyingSheet.Rows.Count);

        // Rendering the wide print area still succeeds and keeps the real content
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(SparseSheetShouldNotMaterializeBlankCellsAcrossWidePrintArea),
            workbook);
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("Body", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SparseSheetShouldRenderImageAnchoredOnBlankRow()
    {
        // Arrange — image anchored far down on a row that carries no cell (a non-materialized row)
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Top";
            // Data in column C keeps the image's column B inside the used range, while row 30 stays empty
            sheet.Cell("C60").Value = "Bottom";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Pic")
                .MoveTo(sheet.Cell("B30"))
                .WithSize(40, 30);
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Row 30 carries no cell, so it is not materialized
        Assert.Null(sheet.UnderlyingSheet.FindCell(30, 2));

        // Act — the image on the blank row is still placed via gap-interpolated offsets
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(SparseSheetShouldRenderImageAnchoredOnBlankRow),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes));
    }

    [Fact]
    public void SparseSheetShouldRenderValueSetOnBlankRow()
    {
        // Arrange — a marker row at the top and a footer defining the used range, blank rows between
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Item}}";
            sheet.Cell("A10").Value = "Footer";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Act — multi-row fill writes into the blank (non-materialized) rows 2 and 3
        var count = sheet.ReplacePlaceholders(new[]
        {
            new Dictionary<string, string?> { ["Item"] = "First" },
            new Dictionary<string, string?> { ["Item"] = "Second" },
            new Dictionary<string, string?> { ["Item"] = "Third" }
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(SparseSheetShouldRenderValueSetOnBlankRow),
            workbook);

        // Assert — values placed on formerly blank rows must be rendered
        Assert.Equal(3, count);
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("First", text, StringComparison.Ordinal);
        Assert.Contains("Second", text, StringComparison.Ordinal);
        Assert.Contains("Third", text, StringComparison.Ordinal);
        Assert.Contains("Footer", text, StringComparison.Ordinal);
    }
}
