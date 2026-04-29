namespace OysterReport.Tests.Internal.OpenXml;

using DocumentFormat.OpenXml.Packaging;

public sealed class SharedStringCatalogTests
{
    [Fact]
    public void LoadShouldReturnEmptyArrayWhenPartIsNull()
    {
        // Act
        var result = SharedStringCatalog.Load(null);

        // Assert
        Assert.Empty(result);
    }

    [Fact]
    public void LoadShouldReturnEmptyArrayWhenWorkbookHasNoSharedStrings()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").SetValue(123);
        });

        // Act
        var catalog = LoadCatalog(stream);

        // Assert
        Assert.Empty(catalog);
    }

    [Fact]
    public void LoadShouldExtractTextFromSimpleSharedStrings()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Alpha";
            sheet.Cell("A2").Value = "Beta";
            sheet.Cell("A3").Value = "Gamma";
        });

        // Act
        var catalog = LoadCatalog(stream);

        // Assert
        Assert.Equal(3, catalog.Length);
        Assert.Contains("Alpha", catalog);
        Assert.Contains("Beta", catalog);
        Assert.Contains("Gamma", catalog);
    }

    [Fact]
    public void LoadShouldConcatenateRichTextRunsIntoSingleString()
    {
        // Arrange — ClosedXML emits rich-text runs when fragments have different formatting.
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var richText = sheet.Cell("A1").GetRichText();
            richText.AddText("Hello").SetBold();
            richText.AddText(" ");
            richText.AddText("World").SetItalic();
        });

        // Act
        var catalog = LoadCatalog(stream);

        // Assert
        Assert.Single(catalog);
        Assert.Equal("Hello World", catalog[0]);
    }

    [Fact]
    public void LoadShouldDeduplicateSharedStringsAcrossSheet()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Repeat";
            sheet.Cell("A2").Value = "Repeat";
            sheet.Cell("A3").Value = "Unique";
        });

        // Act
        var catalog = LoadCatalog(stream);

        // Assert
        Assert.Equal(2, catalog.Length);
    }

    private static string[] LoadCatalog(Stream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        return SharedStringCatalog.Load(doc.WorkbookPart!.SharedStringTablePart);
    }
}
