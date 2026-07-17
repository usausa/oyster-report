namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void MergedCellShouldRenderTextInHorizontalMerge()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "HorizontalMerge";
            sheet.Range("A1:D1").Merge();
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MergedCellShouldRenderTextInHorizontalMerge),
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("HorizontalMerge", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void MergedCellShouldRenderTextInVerticalMerge()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VerticalMerge";
            sheet.Range("A1:A4").Merge();
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MergedCellShouldRenderTextInVerticalMerge),
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("VerticalMerge", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void MergedCellShouldRenderTextInRectangularMerge()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("B2").Value = "RectMerge";
            sheet.Range("B2:D4").Merge();
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MergedCellShouldRenderTextInRectangularMerge),
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("RectMerge", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void MergedCellShouldRenderMultipleMergedRanges()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Range("A1:C1").Merge();
            sheet.Cell("A2").Value = "Left";
            sheet.Range("A2:A4").Merge();
            sheet.Cell("B2").Value = "Right";
            sheet.Range("B2:C4").Merge();
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(MergedCellShouldRenderMultipleMergedRanges), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("Left", text, StringComparison.Ordinal);
        Assert.Contains("Right", text, StringComparison.Ordinal);
    }

    [Fact]
    public void MergedCellShouldNotDuplicateTextFromSubCells()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "MergeOwner";
            sheet.Range("A1:C1").Merge();
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MergedCellShouldNotDuplicateTextFromSubCells),
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var count = CountSubstringOccurrences(TestHelper.ExtractAllText(pdfBytes), "MergeOwner");
        Assert.Equal(1, count);
    }

    [Fact]
    public void MergedCellShouldMoveMergeWhenRowWithHorizontalMergeIsCopied()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Title";
            sheet.Cell("A2").Value = "{{Line}}";
            sheet.Range("A2:C2").Merge();
            sheet.Cell("A3").Value = "Footer";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Line");

        var last = template;
        foreach (var value in new[] { "LineA", "LineB" })
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholder("Line", value);
        }

        template.Delete();

        // Assert — each copied row owns its merged range moved to its own row
        var underlying = sheet.UnderlyingSheet;
        for (var row = 2; row <= 3; row++)
        {
            var owner = underlying.FindCell(row, 1);
            Assert.NotNull(owner);
            Assert.NotNull(owner.Merge);
            Assert.Equal(row, owner.Merge.Range.StartRow);
            Assert.Equal(row, owner.Merge.Range.EndRow);
            Assert.Equal(1, owner.Merge.Range.StartColumn);
            Assert.Equal(3, owner.Merge.Range.EndColumn);
            Assert.Contains(
                underlying.MergedRanges,
                x => (x.Range.StartRow == row) && (x.Range.EndRow == row) && (x.Range.StartColumn == 1) && (x.Range.EndColumn == 3));
        }

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MergedCellShouldMoveMergeWhenRowWithHorizontalMergeIsCopied),
            workbook);
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Equal(1, CountSubstringOccurrences(text, "LineA"));
        Assert.Equal(1, CountSubstringOccurrences(text, "LineB"));
    }

    [Fact]
    public void MergedCellRangeCopyShouldFlattenVerticalMergeAsKnownLimitation()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Group}}";
            sheet.Range("A1:A2").Merge();
            sheet.Cell("B1").Value = "{{First}}";
            sheet.Cell("B2").Value = "{{Second}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var range = sheet.FindRows("Group");
        Assert.Equal(1, range.StartRow);
        Assert.Equal(2, range.EndRow);

        // Act
        range.InsertCopyBelow();

        // Assert — a row-by-row copy cannot reproduce a merge spanning multiple rows;
        // the copied rows are flattened to plain cells (known limitation)
        var underlying = sheet.UnderlyingSheet;
        Assert.NotNull(underlying.FindCell(1, 1)!.Merge);
        Assert.Null(underlying.FindCell(3, 1)!.Merge);
        Assert.Null(underlying.FindCell(4, 1)!.Merge);
        Assert.DoesNotContain(underlying.MergedRanges, x => x.Range.StartRow >= 3);

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(MergedCellRangeCopyShouldFlattenVerticalMergeAsKnownLimitation),
            workbook);
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
    }

    private static int CountSubstringOccurrences(string source, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = source.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }
}
