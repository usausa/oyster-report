namespace OysterReport.Tests;

public sealed partial class ScenarioTests
{
    private const string MultiPageDetailTemplateName = "Template";

    //--------------------------------------------------------------------------------
    // Multi-page detail: page count and footer position
    //--------------------------------------------------------------------------------

    [Theory]
    [InlineData(1, 1)]
    [InlineData(20, 1)]
    [InlineData(21, 2)]
    [InlineData(40, 2)]
    [InlineData(45, 3)]
    public void MultiPageDetailSheetShouldHaveCorrectPageCountAndFooterAtRow21(int totalItems, int expectedPages)
    {
        using var stream = CreateMultiPageDetailTemplate();
        using var workbook = new TemplateWorkbook(stream);
        var items = Enumerable.Range(1, totalItems)
            .Select(static i => (No: i.ToString(CultureInfo.InvariantCulture), Data: $"Item{i}"))
            .ToArray();

        FillMultiPageDetail(workbook, items);

        // The template sheet must be deleted and the correct number of page sheets created
        Assert.Equal(expectedPages, workbook.Sheets.Count);
        Assert.DoesNotContain(workbook.Sheets, static s => s.Name == MultiPageDetailTemplateName);

        // The footer at row 21 must not have moved on any page; row insertions and deletions must not affect it
        foreach (var sheet in workbook.Sheets)
        {
            Assert.Equal("Footer", sheet.GetCellText(21, 1));
        }
    }

    //--------------------------------------------------------------------------------
    // Multi-page detail: PDF content
    //--------------------------------------------------------------------------------

    [Fact]
    public void MultiPageDetailPdfShouldContainAllDataItemsAndFooter()
    {
        const int totalItems = 45;

        using var stream = CreateMultiPageDetailTemplate();
        using var workbook = new TemplateWorkbook(stream);
        var items = Enumerable.Range(1, totalItems)
            .Select(static i => (No: i.ToString(CultureInfo.InvariantCulture), Data: $"Item{i}"))
            .ToArray();

        FillMultiPageDetail(workbook, items);

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MultiPageDetailPdfShouldContainAllDataItemsAndFooter),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);

        // Data from every page must be present (checked with items unique to each page)
        Assert.Contains("Item1", text, StringComparison.Ordinal);    // page 1
        Assert.Contains("Item20", text, StringComparison.Ordinal);   // page 1 last
        Assert.Contains("Item21", text, StringComparison.Ordinal);   // page 2 first
        Assert.Contains("Item40", text, StringComparison.Ordinal);   // page 2 last
        Assert.Contains("Item41", text, StringComparison.Ordinal);   // page 3 first
        Assert.Contains("Item45", text, StringComparison.Ordinal);   // page 3 last (unique: no Item450+)

        // The footer must appear once per page (3 pages = 3 occurrences)
        Assert.Equal(3, CountSubstringOccurrences(text, "Footer"));
    }

    //--------------------------------------------------------------------------------
    // Helpers
    //--------------------------------------------------------------------------------

    // Structure of the template sheet:
    //   Row  1: テンプレート明細行 (Template detail row with {{No}} and {{Data}} placeholders)
    //   Rows 2-21: 空のバッファ行 20行 (20 empty buffer rows)
    //   Row 22: フッタ行 "Footer" (Footer row; shifts to row 21 after the template detail row is deleted)
    private static MemoryStream CreateMultiPageDetailTemplate()
    {
        return TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet(MultiPageDetailTemplateName);
            sheet.Cell("A1").Value = "{{No}}: {{Data}}";
            sheet.Cell("A22").Value = "Footer";
        });
    }

    // 1. Copy the template sheet to create an editing sheet
    // 2. Repeat until the remaining item count or 20 rows is reached
    //    2.1. Insert a copy of the first template row and set data values
    //    2.2. Delete the empty buffer row pushed to row 22 by the insertion to keep the footer at row 22
    // 3. Delete the template detail row (row 1); the footer shifts from row 22 to row 21
    // 4. When more data remains, copy the template again for the next page
    // 5. Finally, delete the original template sheet
    private static void FillMultiPageDetail(
        TemplateWorkbook workbook,
        (string No, string Data)[] items)
    {
        const int rowsPerPage = 20;

        // Footer row in the template. Deleting the buffer row displaced here after each
        // insertion keeps the footer in place; deleting the template row then moves it to row 21
        const int bufferDeleteRow = 22;

        var remaining = items.Length;
        var pageIndex = 0;
        var offset = 0;

        while (remaining > 0)
        {
            pageIndex++;
            var pageSheet = workbook.CopySheet(MultiPageDetailTemplateName, $"Page{pageIndex}");
            var templateDetailRow = pageSheet.GetRow(1);
            var last = templateDetailRow;
            var pageCount = Math.Min(remaining, rowsPerPage);

            for (var i = 0; i < pageCount; i++)
            {
                var newRow = templateDetailRow.InsertCopyAfter(last);
                newRow.ReplacePlaceholder("No", items[offset + i].No);
                newRow.ReplacePlaceholder("Data", items[offset + i].Data);
                last = newRow;

                // Delete the empty buffer row displaced to row 22 to restore the footer to row 22
                pageSheet.DeleteRows(bufferDeleteRow, bufferDeleteRow);
            }

            // Deleting the template detail row shifts the footer from row 22 to row 21
            templateDetailRow.Delete();

            offset += pageCount;
            remaining -= pageCount;
        }

        // Delete the original template sheet used as the basis for editing
        workbook.RemoveSheet(MultiPageDetailTemplateName);
    }
}
