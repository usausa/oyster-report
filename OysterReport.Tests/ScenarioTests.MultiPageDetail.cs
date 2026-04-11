namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>
/// 複数ページ明細印刷のシナリオテスト。
/// 1ページ20行・21行目フッタのテンプレートから複数ページを生成し、
/// シート数・フッタ位置・PDF内容を検証する。
/// Scenario tests for multi-page detail printing.
/// Generates multiple pages from a template with 20 detail rows and a footer at row 21,
/// then verifies page count, footer position, and PDF content.
/// </summary>
public sealed partial class ScenarioTests
{
    private const string MultiPageDetailTemplateName = "Template";

    //--------------------------------------------------------------------------------
    // Multi-page detail: page count and footer position
    //--------------------------------------------------------------------------------

    /// <summary>
    /// 件数に応じたページ数が作成されること、テンプレートシートが削除されること、
    /// および各ページの21行目にフッタが残ることを検証する。
    /// Verifies the correct number of pages is created, the template sheet is deleted,
    /// and the footer remains at row 21 on every page.
    /// </summary>
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
            .Select(static i => (No: i.ToString(), Data: $"Item{i}"))
            .ToArray();

        FillMultiPageDetail(workbook, items);

        // テンプレートシートが削除され、件数に応じたページ数が作成されていること。
        // The template sheet must be deleted and the correct number of page sheets created.
        Assert.Equal(expectedPages, workbook.Sheets.Count);
        Assert.DoesNotContain(workbook.Sheets, static s => s.Name == MultiPageDetailTemplateName);

        // 各ページの21行目のフッタは移動していないこと。行の追加と削除の影響を受けていないこと。
        // The footer at row 21 must not have moved on any page; row insertions and deletions must not affect it.
        foreach (var sheet in workbook.Sheets)
        {
            Assert.Equal("Footer", sheet.UnderlyingWorksheet.Cell(21, 1).GetString());
        }
    }

    //--------------------------------------------------------------------------------
    // Multi-page detail: PDF content
    //--------------------------------------------------------------------------------

    /// <summary>
    /// 全データ項目とフッタが PDF に正しく出力されることを検証する。
    /// Verifies that all data items and the footer text appear correctly in the PDF.
    /// </summary>
    [Fact]
    public void MultiPageDetailPdfShouldContainAllDataItemsAndFooter()
    {
        const int totalItems = 45;

        using var stream = CreateMultiPageDetailTemplate();
        using var workbook = new TemplateWorkbook(stream);
        var items = Enumerable.Range(1, totalItems)
            .Select(static i => (No: i.ToString(), Data: $"Item{i}"))
            .ToArray();

        FillMultiPageDetail(workbook, items);

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MultiPageDetailPdfShouldContainAllDataItemsAndFooter),
            workbook);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);

        // 全ページにデータが含まれていること（各ページに固有の項目で代表して確認）。
        // Data from every page must be present (checked with items unique to each page).
        Assert.Contains("Item1", text, StringComparison.Ordinal);    // page 1
        Assert.Contains("Item20", text, StringComparison.Ordinal);   // page 1 last
        Assert.Contains("Item21", text, StringComparison.Ordinal);   // page 2 first
        Assert.Contains("Item40", text, StringComparison.Ordinal);   // page 2 last
        Assert.Contains("Item41", text, StringComparison.Ordinal);   // page 3 first
        Assert.Contains("Item45", text, StringComparison.Ordinal);   // page 3 last (unique: no Item450+)

        // フッタが全ページ (3ページ) に出力されること。
        // The footer must appear once per page (3 pages = 3 occurrences).
        Assert.Equal(3, CountSubstringOccurrences(text, "Footer"));
    }

    //--------------------------------------------------------------------------------
    // Helpers
    //--------------------------------------------------------------------------------

    /// <summary>
    /// テンプレートワークブックを作成する。
    /// Creates the template workbook used by multi-page detail tests.
    /// </summary>
    /// <remarks>
    /// テンプレートシートの構造:
    /// Structure of the template sheet:
    ///   Row  1: テンプレート明細行 (Template detail row with {{No}} and {{Data}} placeholders)
    ///   Rows 2-21: 空のバッファ行 20行 (20 empty buffer rows)
    ///   Row 22: フッタ行 "Footer" (Footer row; shifts to row 21 after the template detail row is deleted)
    /// </remarks>
    private static MemoryStream CreateMultiPageDetailTemplate()
    {
        return TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet(MultiPageDetailTemplateName);
            sheet.Cell("A1").Value = "{{No}}: {{Data}}";
            sheet.Cell("A22").Value = "Footer";
        });
    }

    /// <summary>
    /// 複数ページ明細印刷のアルゴリズムを実行する。
    /// Executes the multi-page detail printing algorithm.
    /// </summary>
    /// <remarks>
    /// アルゴリズムの概要 / Algorithm overview:
    /// 1. テンプレートシートをコピーして編集用シートを作成する。
    ///    Copy the template sheet to create an editing sheet.
    /// 2. 残データ件数 or 20行になるまで繰り返す。
    ///    Repeat until the remaining item count or 20 rows is reached.
    ///    a. 先頭のテンプレート行を元に行を追加し、データを設定する。
    ///       Insert a copy of the first template row and set data values.
    ///    b. 挿入によって行 22 にずれた空バッファ行を削除し、フッタを行 22 に保持する。
    ///       Delete the empty buffer row pushed to row 22 by the insertion to keep the footer at row 22.
    /// 3. テンプレート明細行 (行 1) を削除するとフッタが行 22 → 行 21 に移動する。
    ///    Delete the template detail row (row 1); the footer shifts from row 22 to row 21.
    /// 4. 行数が足りなくなったら次のシートのコピーを作成して繰り返す。
    ///    When more data remains, copy the template again for the next page.
    /// 5. 最後に最初のテンプレートシートを削除する。
    ///    Finally, delete the original template sheet.
    /// </remarks>
    private static void FillMultiPageDetail(
        TemplateWorkbook workbook,
        IReadOnlyList<(string No, string Data)> items)
    {
        const int rowsPerPage = 20;

        // テンプレートでのフッタ行番号。各挿入後にここにずれた空バッファ行を削除することで
        // フッタを一定位置に保ち、テンプレート行削除後に行 21 へ移動させる。
        // Footer row in the template. Deleting the buffer row displaced here after each
        // insertion keeps the footer in place; deleting the template row then moves it to row 21.
        const int bufferDeleteRow = 22;

        var remaining = items.Count;
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

                // 挿入によって行 22 にずれた空バッファ行を削除し、フッタを行 22 に戻す。
                // Delete the empty buffer row displaced to row 22 to restore the footer to row 22.
                pageSheet.DeleteRows(bufferDeleteRow, bufferDeleteRow);
            }

            // テンプレート明細行を削除することでフッタが行 22 → 行 21 に移動する。
            // Deleting the template detail row shifts the footer from row 22 to row 21.
            templateDetailRow.Delete();

            offset += pageCount;
            remaining -= pageCount;
        }

        // 編集の基となったテンプレートシートを削除する。
        // Delete the original template sheet used as the basis for editing.
        workbook.RemoveSheet(MultiPageDetailTemplateName);
    }
}
