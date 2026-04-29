namespace OysterReport.Tests.Internal.OpenXml;

public sealed class TableLoaderTests
{
    //--------------------------------------------------------------------------------
    // Range parsing
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldParseRangeFromTableReference()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("B2").Value = "Name";
            sheet.Cell("C2").Value = "Value";
            sheet.Cell("B3").Value = "Foo";
            sheet.Cell("C3").Value = 1;
            sheet.Cell("B4").Value = "Bar";
            sheet.Cell("C4").Value = 2;
            sheet.Range("B2:C4").CreateTable("MyTable");
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.Equal(2, table.Range.StartRow);
        Assert.Equal(2, table.Range.StartColumn);
        Assert.Equal(4, table.Range.EndRow);
        Assert.Equal(3, table.Range.EndColumn);
    }

    //--------------------------------------------------------------------------------
    // ShowRowStripes / theme
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldResolveStripeColorWhenThemeAndStripesEnabled()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
            var table = sheet.Tables.First();
            table.Theme = XLTableTheme.TableStyleMedium2;
            table.ShowRowStripes = true;
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.True(table.ShowRowStripes);
        Assert.NotEqual(string.Empty, table.StripeColorHex);
        Assert.StartsWith("#", table.StripeColorHex, StringComparison.Ordinal);
        Assert.Equal(9, table.StripeColorHex.Length);
    }

    [Fact]
    public void LoadShouldLeaveStripeColorEmptyWhenStripesDisabled()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
            var table = sheet.Tables.First();
            table.Theme = XLTableTheme.TableStyleMedium2;
            table.ShowRowStripes = false;
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.False(table.ShowRowStripes);
        Assert.Equal(string.Empty, table.StripeColorHex);
    }

    [Fact]
    public void LoadShouldLeaveStripeColorEmptyWhenThemeNotInStripeBandMap()
    {
        // Arrange — TableStyleLight1 / TableStyleMedium1 are not in StripeBandByStyleName.
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
            var table = sheet.Tables.First();
            table.Theme = XLTableTheme.TableStyleLight1;
            table.ShowRowStripes = true;
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.True(table.ShowRowStripes);
        Assert.Equal(string.Empty, table.StripeColorHex);
    }

    [Theory]
    [InlineData("TableStyleLight2")]
    [InlineData("TableStyleLight9")]
    [InlineData("TableStyleMedium2")]
    [InlineData("TableStyleMedium9")]
    [InlineData("TableStyleMedium16")]
    [InlineData("TableStyleMedium23")]
    public void LoadShouldResolveStripeColorAcrossThemeStyles(string themeName)
    {
        // Arrange
        var themeField = typeof(XLTableTheme).GetField(themeName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        Assert.NotNull(themeField);
        var theme = (XLTableTheme)themeField.GetValue(null)!;

        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
            var table = sheet.Tables.First();
            table.Theme = theme;
            table.ShowRowStripes = true;
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.NotEqual(string.Empty, table.StripeColorHex);
        Assert.Equal(9, table.StripeColorHex.Length);
    }

    //--------------------------------------------------------------------------------
    // Header / totals rows
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldDetectShowHeaderTrueByDefault()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.True(table.ShowHeader);
    }

    [Fact]
    public void LoadShouldDetectShowHeaderFalseWhenHidden()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
            sheet.Tables.First().ShowHeaderRow = false;
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.False(table.ShowHeader);
    }

    [Fact]
    public void LoadShouldDetectShowTotalsTrueWhenEnabled()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
            sheet.Tables.First().ShowTotalsRow = true;
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.True(table.ShowTotals);
    }

    [Fact]
    public void LoadShouldDetectShowTotalsFalseByDefault()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var table = Assert.Single(template.ReportWorkbook.Sheets[0].Tables);

        // Assert
        Assert.False(table.ShowTotals);
    }

    //--------------------------------------------------------------------------------
    // Multiple tables / no tables
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldReturnEmptyTablesWhenWorkbookHasNone()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
        });

        // Act
        using var template = new TemplateWorkbook(stream);

        // Assert
        Assert.Empty(template.ReportWorkbook.Sheets[0].Tables);
    }

    [Fact]
    public void LoadShouldReturnAllTablesWhenSheetHasMultiple()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");

            sheet.Cell("A1").Value = "Name";
            sheet.Cell("B1").Value = "Value";
            sheet.Cell("A2").Value = "Foo";
            sheet.Cell("B2").Value = 1;
            sheet.Range("A1:B2").CreateTable("FirstTable");

            sheet.Cell("D1").Value = "Code";
            sheet.Cell("E1").Value = "Score";
            sheet.Cell("D2").Value = "X";
            sheet.Cell("E2").Value = 99;
            sheet.Range("D1:E2").CreateTable("SecondTable");
        });

        // Act
        using var template = new TemplateWorkbook(stream);
        var tables = template.ReportWorkbook.Sheets[0].Tables;

        // Assert
        Assert.Equal(2, tables.Count);
        Assert.Contains(tables, t => t.Range is { StartColumn: 1, EndColumn: 2 });
        Assert.Contains(tables, t => t.Range is { StartColumn: 4, EndColumn: 5 });
    }

    //--------------------------------------------------------------------------------
    // PDF generation including a styled table should still produce a valid PDF
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldProduceValidPdfForStripedTable()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            BuildSimpleTable(sheet, "A1:B3");
            var table = sheet.Tables.First();
            table.Theme = XLTableTheme.TableStyleMedium2;
            table.ShowRowStripes = true;
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(LoadShouldProduceValidPdfForStripedTable), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("Foo", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    //--------------------------------------------------------------------------------
    // Helpers
    //--------------------------------------------------------------------------------

    private static void BuildSimpleTable(IXLWorksheet sheet, string range)
    {
        var rangeRef = sheet.Range(range);
        var startRow = rangeRef.RangeAddress.FirstAddress.RowNumber;
        var startCol = rangeRef.RangeAddress.FirstAddress.ColumnNumber;

        sheet.Cell(startRow, startCol).Value = "Name";
        sheet.Cell(startRow, startCol + 1).Value = "Value";
        sheet.Cell(startRow + 1, startCol).Value = "Foo";
        sheet.Cell(startRow + 1, startCol + 1).Value = 1;
        sheet.Cell(startRow + 2, startCol).Value = "Bar";
        sheet.Cell(startRow + 2, startCol + 1).Value = 2;
        rangeRef.CreateTable();
    }
}
