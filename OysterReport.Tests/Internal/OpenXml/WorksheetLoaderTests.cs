namespace OysterReport.Tests.Internal.OpenXml;

public sealed class WorksheetLoaderTests
{
    //--------------------------------------------------------------------------------
    // Style instance sharing
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldShareStyleInstancesBetweenCellsWithSameStyle()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "One";
            sheet.Cell("B2").Value = "Two";
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var sheet = template.ReportWorkbook.Sheets[0];

        // Assert — styles are immutable, so cells with the same style index share one instance
        var first = sheet.FindCell(1, 1);
        var second = sheet.FindCell(2, 2);
        Assert.NotNull(first);
        Assert.NotNull(second);
        Assert.Same(first.Style, second.Style);
    }

    [Fact]
    public void CloneShouldKeepStyleInstancesSharedBetweenCells()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "One";
            sheet.Cell("B2").Value = "Two";
        });

        using var template = new TemplateWorkbook(stream);

        // Act
        using var clone = template.Clone();

        // Assert — cloning creates new instances but keeps the sharing between cells
        var original = template.ReportWorkbook.Sheets[0].FindCell(1, 1);
        var first = clone.ReportWorkbook.Sheets[0].FindCell(1, 1);
        var second = clone.ReportWorkbook.Sheets[0].FindCell(2, 2);
        Assert.NotNull(original);
        Assert.NotNull(first);
        Assert.NotNull(second);
        Assert.NotSame(original.Style, first.Style);
        Assert.Same(first.Style, second.Style);
    }
}
