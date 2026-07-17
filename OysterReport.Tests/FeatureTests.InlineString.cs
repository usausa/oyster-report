namespace OysterReport.Tests;

using DocumentFormat.OpenXml;

public sealed partial class FeatureTests
{
    [Fact]
    public void InlineStringShouldJoinRichTextRunsWhenLoading()
    {
        // Arrange — ClosedXML always writes shared strings, so rewrite the cell as a rich inline string
        using var stream = TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                sheet.Cell("A1").Value = "placeholder";
            },
            document =>
            {
                var wsPart = document.WorkbookPart!.WorksheetParts.First();
                var cell = wsPart.Worksheet!.Descendants<Cell>().First(static c => c.CellReference?.Value == "A1");
                cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
                cell.CellValue = null;
                cell.InlineString = new InlineString(
                    new Run(new Text("{{Pla")),
                    new Run(new Text("ceholder}}")));
            });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Act
        var replaced = sheet.ReplacePlaceholder("Placeholder", "Resolved");

        // Assert — the runs are joined so the split placeholder is found and replaced
        Assert.Equal(1, replaced);
        Assert.Equal("Resolved", sheet.GetCellText(1, 1));
    }

    [Fact]
    public void InlineStringShouldLoadPlainTextElement()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                sheet.Cell("A1").Value = "placeholder";
            },
            document =>
            {
                var wsPart = document.WorkbookPart!.WorksheetParts.First();
                var cell = wsPart.Worksheet!.Descendants<Cell>().First(static c => c.CellReference?.Value == "A1");
                cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
                cell.CellValue = null;
                cell.InlineString = new InlineString(new Text("PlainInline"));
            });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Assert
        Assert.Equal("PlainInline", sheet.GetCellText(1, 1));
    }
}
