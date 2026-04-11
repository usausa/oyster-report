namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed class EngineEndToEndTests
{
    [Fact]
    public void GeneratePdfFromFileBasedWorkbookShouldSucceed()
    {
        using var input = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Name}}";
        });

        var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var file = File.Create(tempFile))
            {
                input.CopyTo(file);
            }

            var engine = new OysterReportEngine();
            using var workbook = new TemplateWorkbook(tempFile);
            var sheet = Assert.Single(workbook.Sheets);
            sheet.ReplacePlaceholder("Name", "Bob");

            using var output = new MemoryStream();
            engine.GeneratePdf(workbook, output);

            Assert.True(TestHelper.IsValidPdf(output.ToArray()));
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public void GeneratePdfSingleSheetOverloadShouldSucceed()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Page1").Cell("A1").Value = "FirstPage";
            workbook.AddWorksheet("Page2").Cell("A1").Value = "SecondPage";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var engine = new OysterReportEngine();

        using var output = new MemoryStream();
        engine.GeneratePdf(workbook.Sheets[1], output);

        Assert.True(TestHelper.IsValidPdf(output.ToArray()));
    }
}
