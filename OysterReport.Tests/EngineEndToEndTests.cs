namespace OysterReport.Tests;

using System.IO.Compression;

public sealed class EngineEndToEndTests
{
    [Fact]
    public void GeneratePdfFromFileBasedWorkbookShouldSucceed()
    {
        // Arrange
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

            // Act
            using var output = new MemoryStream();
            engine.GeneratePdf(workbook, output);

            // Assert
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
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Page1").Cell("A1").Value = "FirstPage";
            workbook.AddWorksheet("Page2").Cell("A1").Value = "SecondPage";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var engine = new OysterReportEngine();

        // Act
        using var output = new MemoryStream();
        engine.GeneratePdf(workbook.Sheets[1], output);

        // Assert
        Assert.True(TestHelper.IsValidPdf(output.ToArray()));
    }

    [Fact]
    public void CompressionLevelShouldDefaultToOptimal()
    {
        // Act
        var engine = new OysterReportEngine();

        // Assert
        Assert.Equal(CompressionLevel.Optimal, engine.CompressionLevel);
    }

    [Fact]
    public void CompressionLevelShouldRejectUnknownValue()
    {
        // Arrange
        var engine = new OysterReportEngine();

        // Act / Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => engine.CompressionLevel = (CompressionLevel)99);
    }

    [Theory]
    [InlineData(CompressionLevel.Fastest)]
    [InlineData(CompressionLevel.SmallestSize)]
    [InlineData(CompressionLevel.NoCompression)]
    public void GeneratePdfShouldApplyCompressionLevel(CompressionLevel level)
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            for (var row = 1; row <= 40; row++)
            {
                sheet.Cell(row, 1).Value = $"Line{row}";
                sheet.Cell(row, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
        });

        using var workbook = new TemplateWorkbook(stream);
        var optimal = Generate(new OysterReportEngine { EmbedDocumentMetadata = false }, workbook);

        // Act
        var pdfBytes = Generate(new OysterReportEngine { EmbedDocumentMetadata = false, CompressionLevel = level }, workbook);
        TestHelper.SavePdf($"{nameof(GeneratePdfShouldApplyCompressionLevel)}-{level}", pdfBytes);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("Line40", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
        if (level == CompressionLevel.NoCompression)
        {
            Assert.True(pdfBytes.Length > optimal.Length);
        }
    }

    private static byte[] Generate(OysterReportEngine engine, TemplateWorkbook workbook)
    {
        using var output = new MemoryStream();
        engine.GeneratePdf(workbook, output);
        return output.ToArray();
    }
}
