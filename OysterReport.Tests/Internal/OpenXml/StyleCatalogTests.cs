namespace OysterReport.Tests.Internal.OpenXml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

public sealed class StyleCatalogTests
{
    //--------------------------------------------------------------------------------
    // Defaults when no stylesheet exists
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldReturnDefaultsWhenStylesheetPartMissing()
    {
        // Arrange — build a minimal workbook without a WorkbookStylesPart.
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());
            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1u, Name = "Report" });
        }

        // Act
        var styles = LoadStyles(stream);

        // Assert
        Assert.Single(styles.Fonts);
        Assert.Equal("Calibri", styles.Fonts[0].Name);
        Assert.Equal(11d, styles.Fonts[0].Size);
        Assert.Equal("#FF000000", styles.Fonts[0].ColorHex);
        Assert.Single(styles.Fills);
        Assert.Equal(FillPattern.None, styles.Fills[0].Pattern);
        Assert.Single(styles.Borders);
        Assert.Equal(BorderLineStyle.None, styles.Borders[0].LeftStyle);
        Assert.Single(styles.CellXfs);
        Assert.Empty(styles.CustomNumberFormats);
        Assert.Equal("Calibri", styles.DefaultFontName);
        Assert.Equal(11d, styles.DefaultFontSize);
    }

    //--------------------------------------------------------------------------------
    // Font properties
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldReadBoldItalicUnderlineStrikeFromFontEntry()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "Styled";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Italic = true;
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
            cell.Style.Font.Strikethrough = true;
            cell.Style.Font.FontSize = 14d;
            cell.Style.Font.FontName = "Arial";
        });

        // Act
        var styles = LoadStyles(stream);

        // Assert
        Assert.Contains(styles.Fonts, f =>
            f is { Name: "Arial", Size: 14d, Bold: true, Italic: true, Underline: true, Strike: true });
    }

    [Fact]
    public void LoadShouldUseFirstFontAsDefault()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Report").Cell("A1").Value = "Hello";
        });

        // Act
        var styles = LoadStyles(stream);

        // Assert
        Assert.NotEmpty(styles.Fonts);
        Assert.Equal(styles.Fonts[0].Name, styles.DefaultFontName);
        Assert.Equal(styles.Fonts[0].Size, styles.DefaultFontSize);
    }

    //--------------------------------------------------------------------------------
    // Fills / borders / cell formats
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldReadSolidFillEntry()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "Filled";
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0, 0);
        });

        // Act
        var styles = LoadStyles(stream);

        // Assert — at least one fill entry must be a solid pattern when a cell has background colour.
        Assert.Contains(styles.Fills, f => f.Pattern == FillPattern.Solid);
    }

    [Fact]
    public void LoadShouldReadBorderLineStyles()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "Bordered";
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Medium;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
        });

        // Act
        var styles = LoadStyles(stream);

        // Assert
        Assert.Contains(styles.Borders, b =>
            b is { LeftStyle: BorderLineStyle.Thin, RightStyle: BorderLineStyle.Medium, TopStyle: BorderLineStyle.Thick, BottomStyle: BorderLineStyle.Dashed });
    }

    [Fact]
    public void LoadShouldReadCellAlignmentAndWrapText()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "Aligned";
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            cell.Style.Alignment.WrapText = true;
        });

        // Act
        var styles = LoadStyles(stream);

        // Assert
        Assert.Contains(styles.CellXfs, x =>
            x is { Horizontal: HorizontalAlignment.Center, Vertical: VerticalAlignment.Top, WrapText: true });
    }

    //--------------------------------------------------------------------------------
    // Number formats
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldReadCustomNumberFormat()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = 123.456;
            cell.Style.NumberFormat.Format = "0.000";
        });

        // Act
        var styles = LoadStyles(stream);

        // Assert
        Assert.Contains("0.000", styles.CustomNumberFormats.Values);
    }

    [Theory]
    [InlineData(0, "General")]
    [InlineData(1, "0")]
    [InlineData(2, "0.00")]
    [InlineData(3, "#,##0")]
    [InlineData(4, "#,##0.00")]
    [InlineData(9, "0%")]
    [InlineData(10, "0.00%")]
    [InlineData(14, "mm-dd-yy")]
    [InlineData(22, "m/d/yy h:mm")]
    [InlineData(49, "@")]
    [InlineData(9999, "General")]
    public void ResolveNumberFormatShouldReturnBuiltInCode(int numFmtId, string expected)
    {
        // Arrange
        var styles = LoadEmptyStyles();

        // Act
        var result = styles.ResolveNumberFormat(numFmtId);

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ResolveNumberFormatShouldPreferCustomFormatOverBuiltIn()
    {
        // Arrange — register a custom code for an arbitrary id (>= 164 is convention for custom).
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = 1d;
            cell.Style.NumberFormat.Format = "[Red]0.0";
        });

        var styles = LoadStyles(stream);
        var customId = styles.CustomNumberFormats.First(kv => kv.Value == "[Red]0.0").Key;

        // Act
        var result = styles.ResolveNumberFormat(customId);

        // Assert
        Assert.Equal("[Red]0.0", result);
    }

    [Theory]
    [InlineData(14)]
    [InlineData(15)]
    [InlineData(16)]
    [InlineData(17)]
    [InlineData(18)]
    [InlineData(19)]
    [InlineData(20)]
    [InlineData(21)]
    [InlineData(22)]
    [InlineData(45)]
    [InlineData(46)]
    [InlineData(47)]
    public void IsDateTimeFormatShouldReturnTrueForDateTimeBuiltIns(int numFmtId)
    {
        // Arrange
        var styles = LoadEmptyStyles();

        // Act
        var result = styles.IsDateTimeFormat(numFmtId);

        // Assert
        Assert.True(result);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(9)]
    [InlineData(49)]
    public void IsDateTimeFormatShouldReturnFalseForNonDateTimeBuiltIns(int numFmtId)
    {
        // Arrange
        var styles = LoadEmptyStyles();

        // Act
        var result = styles.IsDateTimeFormat(numFmtId);

        // Assert
        Assert.False(result);
    }

    //--------------------------------------------------------------------------------
    // Helpers
    //--------------------------------------------------------------------------------

    private static StyleCatalog LoadStyles(Stream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        return StyleCatalog.Load(doc.WorkbookPart!);
    }

    private static StyleCatalog LoadEmptyStyles()
    {
        // Build a minimal workbook without a WorkbookStylesPart so the catalog has
        // no custom number-format entries that could shadow built-in ids during lookup.
        var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());
            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1u, Name = "Report" });
        }
        return LoadStyles(stream);
    }
}
