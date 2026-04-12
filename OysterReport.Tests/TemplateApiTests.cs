namespace OysterReport.Tests;

public sealed class TemplateApiTests
{
    [Fact]
    public void ReplacePlaceholderShouldReturnReplacementCount()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{CustomerName}}";
            sheet.Cell("B1").Value = "{{CustomerName}}";
            sheet.Cell("A2").Value = "{{OtherField}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Act
        var customerNameCount = sheet.ReplacePlaceholder("CustomerName", "Alice");
        var otherFieldCount = sheet.ReplacePlaceholder("OtherField", "Value");
        var missingKeyCount = sheet.ReplacePlaceholder("NoSuchKey", "Ignored");

        // Assert
        Assert.Equal(2, customerNameCount);
        Assert.Equal(1, otherFieldCount);
        Assert.Equal(0, missingKeyCount);
    }

    [Fact]
    public void TemplateWorkbookShouldLoadFromFilePath()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Name}}";
        });

        var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var file = File.Create(tempFile))
            {
                stream.CopyTo(file);
            }

            // Act
            using var workbook = new TemplateWorkbook(tempFile);
            var sheet = Assert.Single(workbook.Sheets);
            var count = sheet.ReplacePlaceholder("Name", "Bob");

            // Assert
            Assert.Equal(1, count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
