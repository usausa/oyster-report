namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Fact]
    public void ReplacePlaceholdersShouldReplaceMultipleOnRow()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "name:{{PersonName}} dept:{{PersonDept}} city:{{PersonCity}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.GetRow(1).ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["PersonName"] = "tanaka",
            ["PersonDept"] = "sales",
            ["PersonCity"] = "tokyo"
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldReplaceMultipleOnRow),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("tanaka", text, StringComparison.Ordinal);
        Assert.Contains("sales", text, StringComparison.Ordinal);
        Assert.Contains("tokyo", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersShouldTreatNullValueAsEmpty()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Name: {{Name}}";
            sheet.Cell("B1").Value = "Memo: {{Memo}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).GetRow(1).ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["Name"] = "Alice",
            ["Memo"] = null
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldTreatNullValueAsEmpty),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alice", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Memo}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersShouldReplaceMultipleOnRowRange()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Item: {{ItemName}}";
            sheet.Cell("A2").Value = "Price: {{Price}}";
            sheet.Cell("A3").Value = "Qty: {{Qty}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        Assert.Single(workbook.Sheets).FindRows("ItemName").ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["ItemName"] = "Widget",
            ["Price"] = "980",
            ["Qty"] = "5"
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldReplaceMultipleOnRowRange),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("980", text, StringComparison.Ordinal);
        Assert.Contains("5", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersShouldReplaceAcrossAllSheets()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Cover").Cell("A1").Value = "{{DocTitle}}";
            workbook.AddWorksheet("Body").Cell("A1").Value = "Author: {{Author}}";
            workbook.AddWorksheet("Appendix").Cell("A1").Value = "{{DocTitle}} - Appendix";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        workbook.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["DocTitle"] = "AnnualReport",
            ["Author"] = "Smith"
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldReplaceAcrossAllSheets),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("AnnualReport", text, StringComparison.Ordinal);
        Assert.Contains("Smith", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{DocTitle}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Author}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersSequentialShouldFillConsecutiveRowsFromMarkers()
    {
        // Arrange — markers on row 22 (A22 = {{No}}, B22 = {{Item}}); rows 23+ are empty cells.
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A22").Value = "{{No}}";
            sheet.Cell("B22").Value = "{{Item}}";
            // Pre-allocate empty cells so the underlying model has them available.
            for (var r = 23; r <= 26; r++)
            {
                sheet.Cell($"A{r}").Value = string.Empty;
                sheet.Cell($"B{r}").Value = string.Empty;
            }
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        var rows = new[]
        {
            new Dictionary<string, string?> { ["No"] = "1", ["Item"] = "Alpha" },
            new Dictionary<string, string?> { ["No"] = "2", ["Item"] = "Beta" },
            new Dictionary<string, string?> { ["No"] = "3", ["Item"] = "Gamma" }
        };

        // Act
        var written = sheet.ReplacePlaceholders(rows);

        // Assert
        Assert.Equal(6, written);
        Assert.Equal("1", sheet.GetCellText(22, 1));
        Assert.Equal("Alpha", sheet.GetCellText(22, 2));
        Assert.Equal("2", sheet.GetCellText(23, 1));
        Assert.Equal("Beta", sheet.GetCellText(23, 2));
        Assert.Equal("3", sheet.GetCellText(24, 1));
        Assert.Equal("Gamma", sheet.GetCellText(24, 2));

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersSequentialShouldFillConsecutiveRowsFromMarkers),
            workbook);
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Gamma", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{No}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Item}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacePlaceholdersSequentialShouldCreateCellsWhenMissing()
    {
        // Arrange — only the marker row has cells; rows 23+ are missing entirely.
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A22").Value = "{{Code}}";
            sheet.Cell("C22").Value = "{{Label}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        var rows = new[]
        {
            new Dictionary<string, string?> { ["Code"] = "001", ["Label"] = "Alpha" },
            new Dictionary<string, string?> { ["Code"] = "002", ["Label"] = "Beta" }
        };

        // Act
        sheet.ReplacePlaceholders(rows);

        // Assert
        Assert.Equal("001", sheet.GetCellText(22, 1));
        Assert.Equal("Alpha", sheet.GetCellText(22, 3));
        Assert.Equal("002", sheet.GetCellText(23, 1));
        Assert.Equal("Beta", sheet.GetCellText(23, 3));
    }

    [Fact]
    public void ReplacePlaceholdersSequentialShouldTreatNullValueAsEmpty()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Name}}";
            sheet.Cell("B1").Value = "{{Memo}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        var rows = new[]
        {
            new Dictionary<string, string?> { ["Name"] = "Alice", ["Memo"] = null }
        };

        // Act
        sheet.ReplacePlaceholders(rows);

        // Assert
        Assert.Equal("Alice", sheet.GetCellText(1, 1));
        Assert.Equal(string.Empty, sheet.GetCellText(1, 2));
    }

    [Fact]
    public void ReplacePlaceholdersSequentialShouldIgnoreUnknownMarkers()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Known}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        var rows = new[]
        {
            new Dictionary<string, string?> { ["Known"] = "yes", ["Missing"] = "ignored" }
        };

        // Act
        var written = sheet.ReplacePlaceholders(rows);

        // Assert
        Assert.Equal(1, written);
        Assert.Equal("yes", sheet.GetCellText(1, 1));
    }

    [Fact]
    public void FindMarkerPositionShouldReturnRowAndColumn()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("C5").Value = "{{Target}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Act
        var position = sheet.FindMarkerPosition("Target");
        var found = sheet.TryFindMarkerPosition("Target", out var got);
        var notFound = sheet.TryFindMarkerPosition("Missing", out _);

        // Assert
        Assert.Equal((5, 3), position);
        Assert.True(found);
        Assert.Equal((5, 3), got);
        Assert.False(notFound);
    }

    [Fact]
    public void FindMarkerPositionShouldThrowWhenMissing()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Other}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => sheet.FindMarkerPosition("Missing"));
    }

    [Fact]
    public void SetCellValueShouldOverwriteExistingCellAndCreateMissing()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "old";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        // Act
        sheet.SetCellValue(1, 1, "new");
        sheet.SetCellValue(2, 2, "fresh");
        sheet.SetCellValue(3, 3, null);

        // Assert
        Assert.Equal("new", sheet.GetCellText(1, 1));
        Assert.Equal("fresh", sheet.GetCellText(2, 2));
        Assert.Equal(string.Empty, sheet.GetCellText(3, 3));
    }

    [Fact]
    public void ReplacePlaceholdersShouldWorkWithExpandedRowsInLoop()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "HEADER";
            sheet.Cell("A2").Value = "{{Code}}";
            sheet.Cell("B2").Value = "{{Label}}";
            sheet.Cell("C2").Value = "{{Value}}";
            sheet.Cell("A3").Value = "FOOTER";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Code");

        var items = new[]
        {
            new Dictionary<string, string?> { ["Code"] = "001", ["Label"] = "Alpha", ["Value"] = "100" },
            new Dictionary<string, string?> { ["Code"] = "002", ["Label"] = "Beta",  ["Value"] = "200" },
            new Dictionary<string, string?> { ["Code"] = "003", ["Label"] = "Gamma", ["Value"] = "300" }
        };

        var last = template;
        foreach (var item in items)
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholders(item);
        }

        template.Delete();

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(ReplacePlaceholdersShouldWorkWithExpandedRowsInLoop),
            workbook);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("001", text, StringComparison.Ordinal);
        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("002", text, StringComparison.Ordinal);
        Assert.Contains("Beta", text, StringComparison.Ordinal);
        Assert.Contains("003", text, StringComparison.Ordinal);
        Assert.Contains("Gamma", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Code}}", text, StringComparison.Ordinal);
    }
}
