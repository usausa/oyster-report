namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Theory]
    [InlineData("YellowBg", 255, 255, 0)]
    [InlineData("LightBlueBg", 173, 216, 230)]
    [InlineData("GrayBg", 192, 192, 192)]
    public void FillColorShouldRenderTextOnColoredBackground(string cellValue, int r, int g, int b)
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(r, g, b);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(FillColorShouldRenderTextOnColoredBackground)}_{cellValue}",
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FillColorShouldRenderMultipleDifferentBackgroundColors()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var redCell = sheet.Cell("A1");
            redCell.Value = "RedBg";
            redCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 200, 200);
            var greenCell = sheet.Cell("A2");
            greenCell.Value = "GreenBg";
            greenCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 255, 200);
            var blueCell = sheet.Cell("A3");
            blueCell.Value = "BlueBg";
            blueCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 200, 255);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(FillColorShouldRenderMultipleDifferentBackgroundColors),
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("RedBg", text, StringComparison.Ordinal);
        Assert.Contains("GreenBg", text, StringComparison.Ordinal);
        Assert.Contains("BlueBg", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FillColorShouldRenderThemeBackgroundColor()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeBg";
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.25);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FillColorShouldRenderThemeBackgroundColor), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeBg", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FillColorShouldIgnoreAlphaOfRgbColor()
    {
        // Arrange — openpyxl writes 6-digit colors as 00RRGGBB; Excel paints the fill without looking at the alpha
        using var stream = TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                var zeroAlphaCell = sheet.Cell("A1");
                zeroAlphaCell.Value = "ZeroAlphaBg";
                zeroAlphaCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 0, 0);
                var halfAlphaCell = sheet.Cell("A2");
                halfAlphaCell.Value = "HalfAlphaBg";
                halfAlphaCell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 0, 255);
            },
            document =>
            {
                TestWorkbookFactory.ReplaceRgbColor(document, "FFFF0000", "00FF0000");
                TestWorkbookFactory.ReplaceRgbColor(document, "FF0000FF", "800000FF");
            });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FillColorShouldIgnoreAlphaOfRgbColor), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var fillColors = TestHelper.GetFillColors(pdfBytes);
        Assert.Contains((1d, 0d, 0d), fillColors);
        Assert.Contains((0d, 0d, 1d), fillColors);
        Assert.False(TestHelper.HasTransparency(pdfBytes));
    }

    [Fact]
    public void FillColorShouldPaintSolidFillWithoutForegroundColorBlack()
    {
        // Arrange — openpyxl omits the black foreground color 00000000; Excel paints such a solid fill black and ignores bgColor
        using var stream = TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                var noColorCell = sheet.Cell("A1");
                noColorCell.Value = "NoColorBg";
                noColorCell.Style.Font.FontColor = XLColor.White;
                noColorCell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 0, 0);
                var backgroundOnlyCell = sheet.Cell("A2");
                backgroundOnlyCell.Value = "BackgroundOnlyBg";
                backgroundOnlyCell.Style.Font.FontColor = XLColor.White;
                backgroundOnlyCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);
            },
            document =>
            {
                var fills = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Descendants<PatternFill>().ToList();
                fills.First(static x => x.ForegroundColor?.Rgb?.Value == "FF000000").RemoveAllChildren();
                var backgroundOnly = fills.First(static x => x.ForegroundColor?.Rgb?.Value == "FFFFFF00");
                backgroundOnly.RemoveAllChildren();
                backgroundOnly.BackgroundColor = new BackgroundColor { Rgb = "FFFFFF00" };
            });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FillColorShouldPaintSolidFillWithoutForegroundColorBlack), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var fillColors = TestHelper.GetFillColors(pdfBytes);
        Assert.Equal(2, fillColors.Count(static x => x == (0d, 0d, 0d)));
        Assert.DoesNotContain((1d, 1d, 0d), fillColors);
    }
}
