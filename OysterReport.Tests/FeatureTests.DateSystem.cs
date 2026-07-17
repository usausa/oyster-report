namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void DateSystemShouldShiftDateWhen1904SystemEnabled()
    {
        // Arrange — ClosedXML writes serials in the 1900 system; flipping the flag reinterprets them.
        // Serial 43831 means 2020-01-01 in the 1900 system and 2024-01-02 in the 1904 system.
        using var stream = TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                var cell = sheet.Cell("A1");
                cell.Value = new DateTime(2020, 1, 1);
                cell.Style.DateFormat.Format = "yyyy/mm/dd";
            },
            document =>
            {
                var wb = document.WorkbookPart!.Workbook!;
                wb.WorkbookProperties ??= new WorkbookProperties();
                wb.WorkbookProperties.Date1904 = true;
            });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var cell = sheet.UnderlyingSheet.FindCell(1, 1);

        // Assert
        Assert.NotNull(cell);
        Assert.Equal(CellValueKind.DateTime, cell.Value.Kind);
        Assert.Equal(new DateTime(2024, 1, 2), cell.Value.RawValue);
        Assert.Contains("2024", sheet.GetCellText(1, 1), StringComparison.Ordinal);
    }

    [Fact]
    public void DateSystemShouldKeepDateWhen1900SystemDefault()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = new DateTime(2020, 1, 1);
            cell.Style.DateFormat.Format = "yyyy/mm/dd";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var cell = sheet.UnderlyingSheet.FindCell(1, 1);

        // Assert
        Assert.NotNull(cell);
        Assert.Equal(CellValueKind.DateTime, cell.Value.Kind);
        Assert.Equal(new DateTime(2020, 1, 1), cell.Value.RawValue);
    }
}
