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
