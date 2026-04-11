namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed class TemplateApiTests
{
    [Fact]
    public void ReplacePlaceholderShouldReturnReplacementCount()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{CustomerName}}";
            sheet.Cell("B1").Value = "{{CustomerName}}";
            sheet.Cell("A2").Value = "{{OtherField}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // CustomerName は 2 セルに存在するので 2 が返るはず
        Assert.Equal(2, sheet.ReplacePlaceholder("CustomerName", "Alice"));
        // OtherField は 1 セルにのみ存在
        Assert.Equal(1, sheet.ReplacePlaceholder("OtherField", "Value"));
        // 存在しないキーは 0
        Assert.Equal(0, sheet.ReplacePlaceholder("NoSuchKey", "Ignored"));
    }

    [Fact]
    public void TemplateWorkbookShouldLoadFromFilePath()
    {
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

            using var workbook = new TemplateWorkbook(tempFile);
            var sheet = Assert.Single(workbook.Sheets);
            var count = sheet.ReplacePlaceholder("Name", "Bob");

            Assert.Equal(1, count);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
