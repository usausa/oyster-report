namespace OysterReport.Tests.Internal.OpenXml;

public sealed class OpenXmlLoaderTests
{
    //--------------------------------------------------------------------------------
    // Metadata
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldUseWorkbookTitleAsTemplateName()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.Properties.Title = "MyTemplate";
            workbook.AddWorksheet("Report").Cell("A1").Value = "Hello";
        });

        // Act
        using var template = new TemplateWorkbook(stream);

        // Assert
        Assert.Equal("MyTemplate", template.ReportWorkbook.Metadata.TemplateName);
    }

    [Fact]
    public void LoadShouldUseDefaultTemplateNameWhenTitleIsMissing()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Report").Cell("A1").Value = "Hello";
        });

        // Act
        using var template = new TemplateWorkbook(stream);

        // Assert
        Assert.Equal("Workbook", template.ReportWorkbook.Metadata.TemplateName);
    }

    //--------------------------------------------------------------------------------
    // Print area parsing
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldParsePrintAreaForRange()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Cell";
            sheet.PageSetup.PrintAreas.Add("A1:C5");
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var printArea = template.ReportWorkbook.Sheets[0].PrintArea;

        // Assert
        Assert.NotNull(printArea);
        Assert.Equal(1, printArea.Range.StartRow);
        Assert.Equal(1, printArea.Range.StartColumn);
        Assert.Equal(5, printArea.Range.EndRow);
        Assert.Equal(3, printArea.Range.EndColumn);
    }

    [Fact]
    public void LoadShouldParsePrintAreaForSingleCell()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("B3").Value = "Only";
            sheet.PageSetup.PrintAreas.Add("B3");
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var printArea = template.ReportWorkbook.Sheets[0].PrintArea;

        // Assert
        Assert.NotNull(printArea);
        Assert.Equal(3, printArea.Range.StartRow);
        Assert.Equal(2, printArea.Range.StartColumn);
        Assert.Equal(3, printArea.Range.EndRow);
        Assert.Equal(2, printArea.Range.EndColumn);
    }

    [Fact]
    public void LoadShouldLeavePrintAreaNullWhenNotDefined()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Cell";
        });

        // Act
        using var template = new TemplateWorkbook(stream);

        // Assert
        Assert.Null(template.ReportWorkbook.Sheets[0].PrintArea);
    }

    [Fact]
    public void LoadShouldAssignPrintAreaToCorrectSheetUsingLocalSheetId()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var first = workbook.AddWorksheet("First");
            first.Cell("A1").Value = "First";

            var second = workbook.AddWorksheet("Second");
            second.Cell("A1").Value = "Second";
            second.PageSetup.PrintAreas.Add("A1:B2");
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var sheets = template.ReportWorkbook.Sheets;

        // Assert
        Assert.Null(sheets[0].PrintArea);
        var printArea = sheets[1].PrintArea;
        Assert.NotNull(printArea);
        Assert.Equal(2, printArea.Range.EndRow);
        Assert.Equal(2, printArea.Range.EndColumn);
    }

    //--------------------------------------------------------------------------------
    // Stream vs path overload
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldProduceEquivalentResultFromPathAndStream()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "FromStreamOrPath";
            sheet.PageSetup.PrintAreas.Add("A1:B2");
        });

        var tempPath = Path.Combine(Path.GetTempPath(), $"oyster_{Guid.NewGuid():N}.xlsx");
        try
        {
            File.WriteAllBytes(tempPath, stream.ToArray());

            // Act
            using var fromStream = new TemplateWorkbook(stream);
            using var fromPath = new TemplateWorkbook(tempPath);

            // Assert
            Assert.Equal(fromStream.ReportWorkbook.Sheets.Count, fromPath.ReportWorkbook.Sheets.Count);
            Assert.Equal(
                fromStream.ReportWorkbook.Sheets[0].PrintArea?.Range,
                fromPath.ReportWorkbook.Sheets[0].PrintArea?.Range);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }
}
