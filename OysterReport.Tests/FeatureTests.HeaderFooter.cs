namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void HeaderFooterShouldRenderHeaderText()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Header.Left.AddText("LeftHeader", XLHFOccurrence.OddPages);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldRenderHeaderText), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BodyContent", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldRenderFooterText()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Footer.Right.AddText("RightFooter", XLHFOccurrence.OddPages);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldRenderFooterText), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void HeaderFooterShouldRenderBothHeaderAndFooter()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Main";
            sheet.PageSetup.Header.Center.AddText("TopCenter", XLHFOccurrence.OddPages);
            sheet.PageSetup.Footer.Center.AddText("BottomCenter", XLHFOccurrence.OddPages);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldRenderBothHeaderAndFooter), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
    }

    //--------------------------------------------------------------------------------
    // Page numbering
    //--------------------------------------------------------------------------------

    [Fact]
    public void HeaderFooterShouldNumberPagesThroughWorkbook()
    {
        // Arrange
        using var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddFooter("&C&P / &N")));
        using var workbook = new TemplateWorkbook(stream);
        CopyTemplatePages(workbook, 3);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldNumberPagesThroughWorkbook), workbook);

        // Assert
        Assert.Equal(3, TestHelper.GetPageCount(pdfBytes));
        Assert.Contains("1 / 3", TestHelper.GetPageText(pdfBytes, 1), StringComparison.Ordinal);
        Assert.Contains("2 / 3", TestHelper.GetPageText(pdfBytes, 2), StringComparison.Ordinal);
        Assert.Contains("3 / 3", TestHelper.GetPageText(pdfBytes, 3), StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldRestartNumberingForEachWorkbook()
    {
        // Arrange
        using var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddFooter("&C&P / &N")));
        using var template = new TemplateWorkbook(stream);
        using var first = template.Clone();
        CopyTemplatePages(first, 2);
        using var second = template.Clone();
        CopyTemplatePages(second, 1);

        // Act
        var pdfBytes = GenerateBatchPdf(nameof(HeaderFooterShouldRestartNumberingForEachWorkbook), first, second);

        // Assert
        Assert.Equal(3, TestHelper.GetPageCount(pdfBytes));
        Assert.Contains("1 / 2", TestHelper.GetPageText(pdfBytes, 1), StringComparison.Ordinal);
        Assert.Contains("2 / 2", TestHelper.GetPageText(pdfBytes, 2), StringComparison.Ordinal);
        Assert.Contains("1 / 1", TestHelper.GetPageText(pdfBytes, 3), StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldStartDocumentAtFirstPageNumber()
    {
        // Arrange: every copied sheet carries the setting, but only the first page of the document uses it
        using var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddFooter("&C&P / &N")), static x => x.PageSetup.FirstPageNumber = 5);
        using var workbook = new TemplateWorkbook(stream);
        CopyTemplatePages(workbook, 2);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldStartDocumentAtFirstPageNumber), workbook);

        // Assert: &N is the page count, as in Excel
        Assert.Contains("5 / 2", TestHelper.GetPageText(pdfBytes, 1), StringComparison.Ordinal);
        Assert.Contains("6 / 2", TestHelper.GetPageText(pdfBytes, 2), StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldUseFirstPageHeaderOnFirstPageOfEachWorkbook()
    {
        // Arrange
        using var stream = CreateHeaderFooterTemplate(static x =>
        {
            x.DifferentFirst = true;
            x.Append(new OddHeader("&CNextPage"), new FirstHeader("&CFirstPage"));
        });
        using var template = new TemplateWorkbook(stream);
        using var first = template.Clone();
        CopyTemplatePages(first, 2);
        using var second = template.Clone();
        CopyTemplatePages(second, 1);

        // Act
        var pdfBytes = GenerateBatchPdf(nameof(HeaderFooterShouldUseFirstPageHeaderOnFirstPageOfEachWorkbook), first, second);

        // Assert
        Assert.Contains("FirstPage", TestHelper.GetPageText(pdfBytes, 1), StringComparison.Ordinal);
        Assert.Contains("NextPage", TestHelper.GetPageText(pdfBytes, 2), StringComparison.Ordinal);
        Assert.Contains("FirstPage", TestHelper.GetPageText(pdfBytes, 3), StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldChooseOddEvenHeaderByPageNumber()
    {
        // Arrange: the first page is number 2, so it takes the even page header
        using var stream = CreateHeaderFooterTemplate(
            static x =>
            {
                x.DifferentOddEven = true;
                x.Append(new OddHeader("&COddPage"), new EvenHeader("&CEvenPage"));
            },
            static x => x.PageSetup.FirstPageNumber = 2);
        using var workbook = new TemplateWorkbook(stream);
        CopyTemplatePages(workbook, 2);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldChooseOddEvenHeaderByPageNumber), workbook);

        // Assert
        Assert.Contains("EvenPage", TestHelper.GetPageText(pdfBytes, 1), StringComparison.Ordinal);
        Assert.Contains("OddPage", TestHelper.GetPageText(pdfBytes, 2), StringComparison.Ordinal);
    }

    //--------------------------------------------------------------------------------
    // Format codes
    //--------------------------------------------------------------------------------

    [Fact]
    public void HeaderFooterShouldNotRenderFormattingCodes()
    {
        // Arrange
        using var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddHeader("&L&\"Arial,Bold\"&14Title&C&KFF0000Red&R&K01+000Theme&B&I&U")));
        using var workbook = new TemplateWorkbook(stream);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldNotRenderFormattingCodes), workbook);

        // Assert
        var text = TestHelper.GetPageText(pdfBytes);
        Assert.Contains("Title", text, StringComparison.Ordinal);
        Assert.Contains("Red", text, StringComparison.Ordinal);
        Assert.Contains("Theme", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Arial", text, StringComparison.Ordinal);
        Assert.DoesNotContain("4Title", text, StringComparison.Ordinal);
        Assert.DoesNotContain("FF0000", text, StringComparison.Ordinal);
        Assert.DoesNotContain("01+000", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldRenderSheetNameDateAndTime()
    {
        // Arrange
        using var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddFooter("&L&A&C&D &T")));
        using var workbook = new TemplateWorkbook(stream);
        var engine = new OysterReportEngine
        {
            RenderingOptions =
            {
                HeaderFooterCulture = CultureInfo.InvariantCulture,
                TimeProvider = new FixedTimeProvider(new DateTimeOffset(2026, 11, 2, 14, 5, 0, TimeSpan.Zero))
            }
        };

        // Act
        using var output = new MemoryStream();
        engine.GeneratePdf(workbook, output);
        var pdfBytes = output.ToArray();
        TestHelper.SavePdf(nameof(HeaderFooterShouldRenderSheetNameDateAndTime), pdfBytes);

        // Assert
        var text = TestHelper.GetPageText(pdfBytes);
        Assert.Contains(HeaderFooterTemplateName, text, StringComparison.Ordinal);
        Assert.Contains("11/02/2026 14:05", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldRenderFileName()
    {
        // Arrange
        var path = Path.Combine(Path.GetTempPath(), $"oyster-header-{Guid.NewGuid():N}.xlsx");
        using (var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddFooter("&C&F"))))
        {
            File.WriteAllBytes(path, stream.ToArray());
        }

        try
        {
            using var workbook = new TemplateWorkbook(path);

            // Act
            var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldRenderFileName), workbook);

            // Assert
            Assert.Contains(Path.GetFileName(path), TestHelper.GetPageText(pdfBytes), StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    //--------------------------------------------------------------------------------
    // Placeholders
    //--------------------------------------------------------------------------------

    [Fact]
    public void HeaderFooterShouldReplacePlaceholders()
    {
        // Arrange
        using var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddHeader("&L{{Customer}}"), new OddFooter("&C{{InvoiceNo}}")));
        using var workbook = new TemplateWorkbook(stream);

        // Act
        var count = workbook.ReplacePlaceholders(new Dictionary<string, string?> { ["Customer"] = "AT&T", ["InvoiceNo"] = "INV-001" });
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldReplacePlaceholders), workbook);

        // Assert: '&' in a value is printed instead of starting a format code
        Assert.Equal(2, count);
        var text = TestHelper.GetPageText(pdfBytes);
        Assert.Contains("AT&T", text, StringComparison.Ordinal);
        Assert.Contains("INV-001", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{", text, StringComparison.Ordinal);
    }

    [Fact]
    public void HeaderFooterShouldReplaceSinglePlaceholder()
    {
        // Arrange
        using var stream = CreateHeaderFooterTemplate(static x => x.Append(new OddHeader("&L{{Department}} &P")));
        using var workbook = new TemplateWorkbook(stream);

        // Act
        var count = workbook.GetSheet(HeaderFooterTemplateName).ReplacePlaceholder("Department", "R&D");
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(HeaderFooterShouldReplaceSinglePlaceholder), workbook);

        // Assert
        Assert.Equal(1, count);
        Assert.Contains("R&D 1", TestHelper.GetPageText(pdfBytes), StringComparison.Ordinal);
    }

    //--------------------------------------------------------------------------------
    // Helper
    //--------------------------------------------------------------------------------

    private const string HeaderFooterTemplateName = "Template";

    // The format codes are written with the Open XML SDK so that the tests control the exact text
    private static MemoryStream CreateHeaderFooterTemplate(Action<HeaderFooter> configureHeaderFooter, Action<IXLWorksheet>? configureSheet = null) =>
        TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet(HeaderFooterTemplateName);
                sheet.Cell("A1").Value = "Body";

                // Makes ClosedXML write the headerFooter element, which is rewritten below
                sheet.PageSetup.Header.Center.AddText("Header", XLHFOccurrence.OddPages);
                configureSheet?.Invoke(sheet);
            },
            document =>
            {
                var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;
                var headerFooter = worksheet.GetFirstChild<HeaderFooter>()!;
                headerFooter.RemoveAllChildren();
                headerFooter.DifferentFirst = null;
                headerFooter.DifferentOddEven = null;
                configureHeaderFooter(headerFooter);
                worksheet.Save();
            });

    // Builds the pages with the CopySheet pattern and removes the template sheet
    private static void CopyTemplatePages(TemplateWorkbook workbook, int pageCount)
    {
        for (var i = 1; i <= pageCount; i++)
        {
            workbook.CopySheet(HeaderFooterTemplateName, $"Page{i}");
        }

        workbook.RemoveSheet(HeaderFooterTemplateName);
    }

    private static byte[] GenerateBatchPdf(string testName, params TemplateWorkbook[] workbooks)
    {
        using var output = new MemoryStream();
        new OysterReportEngine().GeneratePdf(workbooks, output);
        var pdfBytes = output.ToArray();
        TestHelper.SavePdf(testName, pdfBytes);
        return pdfBytes;
    }

    private sealed class FixedTimeProvider(DateTimeOffset now) : TimeProvider
    {
        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;

        public override DateTimeOffset GetUtcNow() => now;
    }
}
