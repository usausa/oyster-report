namespace OysterReport.Tests;

public sealed partial class FeatureTests
{
    [Theory]
    [InlineData(XLBorderStyleValues.Thin, "ThinBorder")]
    [InlineData(XLBorderStyleValues.Medium, "MediumBorder")]
    [InlineData(XLBorderStyleValues.Thick, "ThickBorder")]
    [InlineData(XLBorderStyleValues.Double, "DoubleBorder")]
    [InlineData(XLBorderStyleValues.Dashed, "DashedBorder")]
    [InlineData(XLBorderStyleValues.Dotted, "DottedBorder")]
    public void BorderShouldRenderCellWithBorder(XLBorderStyleValues borderStyle, string cellValue)
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = cellValue;
            cell.Style.Border.OutsideBorder = borderStyle;
            cell.Style.Border.OutsideBorderColor = XLColor.Black;
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(
            $"{nameof(BorderShouldRenderCellWithBorder)}_{borderStyle}",
            stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void BorderShouldRenderColoredBorder()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = "ColoredBorder";
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            cell.Style.Border.OutsideBorderColor = XLColor.FromArgb(255, 0, 0);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(BorderShouldRenderColoredBorder), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ColoredBorder", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void BorderShouldIgnoreAlphaOfRgbColor()
    {
        // Arrange — openpyxl writes 6-digit colors as 00RRGGBB; Excel draws the border without looking at the alpha
        using var stream = TestWorkbookFactory.CreateWorkbook(
            workbook =>
            {
                var sheet = workbook.AddWorksheet("Report");
                var cell = sheet.Cell("B2");
                cell.Value = "ZeroAlphaBorder";
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                cell.Style.Border.OutsideBorderColor = XLColor.FromArgb(0, 255, 0);
            },
            document => TestWorkbookFactory.ReplaceRgbColor(document, "FF00FF00", "0000FF00"));

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(BorderShouldIgnoreAlphaOfRgbColor), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var fillColors = TestHelper.GetFillColors(pdfBytes, 1);
        Assert.Equal(4, fillColors.Count(static x => x == (0d, 1d, 0d)));
        Assert.DoesNotContain((0d, 0d, 0d), fillColors);
    }

    [Fact]
    public void BorderShouldRenderTableWithAllSideBorders()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            for (var row = 1; row <= 3; row++)
            {
                for (var col = 1; col <= 3; col++)
                {
                    var cell = sheet.Cell(row, col);
                    cell.Value = $"R{row}C{col}";
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(BorderShouldRenderTableWithAllSideBorders), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("R1C1", text, StringComparison.Ordinal);
        Assert.Contains("R3C3", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(XLBorderStyleValues.Thin)]
    [InlineData(XLBorderStyleValues.Medium)]
    [InlineData(XLBorderStyleValues.Thick)]
    public void BorderShouldFillCornersOfFrame(XLBorderStyleValues borderStyle)
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var range = sheet.Range("B2:C3");
            foreach (var cell in range.Cells())
            {
                cell.Value = "Frame";
            }

            range.Style.Border.OutsideBorder = borderStyle;
            range.Style.Border.OutsideBorderColor = XLColor.Black;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorderColor = XLColor.Black;
        });
        using var workbook = new TemplateWorkbook(stream);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave($"{nameof(BorderShouldFillCornersOfFrame)}_{borderStyle}", workbook);

        // Assert: the corners are filled up to the outer edge of the frame, and nothing goes past it
        var options = new ReportRenderOption();
        var width = borderStyle switch
        {
            XLBorderStyleValues.Thick => options.ThickBorderWidth,
            XLBorderStyleValues.Medium => options.MediumBorderWidth,
            _ => options.NormalBorderWidth
        };
        var half = width / 2;
        var topLeft = GetCellBounds(workbook, "B2");
        var bottomRight = GetCellBounds(workbook, "C3");
        var rects = TestHelper.GetFilledRects(pdfBytes, 1, (0d, 0d, 0d));
        Assert.True(IsFilled(rects, topLeft.X - (half / 2), topLeft.Y - (half / 2)));
        Assert.True(IsFilled(rects, bottomRight.Right + (half / 2), topLeft.Y - (half / 2)));
        Assert.True(IsFilled(rects, bottomRight.Right + (half / 2), bottomRight.Bottom + (half / 2)));
        Assert.True(IsFilled(rects, topLeft.X - (half / 2), bottomRight.Bottom + (half / 2)));
        Assert.False(IsFilled(rects, topLeft.X - half - 0.1, topLeft.Y - (half / 2)));
        Assert.False(IsFilled(rects, topLeft.X - (half / 2), topLeft.Y - half - 0.1));
        Assert.False(IsFilled(rects, topLeft.Right, topLeft.Y - half - 0.1));
    }

    [Fact]
    public void BorderShouldJoinDoubleFrameAtCorners()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = "Double";
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Double;
            cell.Style.Border.OutsideBorderColor = XLColor.Black;
        });
        using var workbook = new TemplateWorkbook(stream);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(BorderShouldJoinDoubleFrameAtCorners), workbook);

        // Assert: outer strokes meet at the outer corner, inner strokes at the inner corner, and the gap stays open
        var width = new ReportRenderOption().NormalBorderWidth;
        var offset = Math.Max(RenderConstants.MinimumDoubleBorderGap, width * RenderConstants.DoubleBorderGapWidthMultiplier) / 2;
        var bounds = GetCellBounds(workbook, "B2");
        var rects = TestHelper.GetFilledRects(pdfBytes, 1, (0d, 0d, 0d));
        Assert.True(IsFilled(rects, bounds.X - offset, bounds.Y - offset));
        Assert.True(IsFilled(rects, bounds.X + offset, bounds.Y + offset));
        Assert.False(IsFilled(rects, bounds.X, bounds.Y));
        Assert.False(IsFilled(rects, bounds.X + 0.1, bounds.Y + offset));
        Assert.False(IsFilled(rects, bounds.X + offset, bounds.Y + 0.1));
    }

    [Fact]
    public void BorderShouldKeepFreeLineEndsAtCellEdges()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = "Underline";
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.BottomBorderColor = XLColor.Black;
        });
        using var workbook = new TemplateWorkbook(stream);

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(BorderShouldKeepFreeLineEndsAtCellEdges), workbook);

        // Assert
        var bounds = GetCellBounds(workbook, "B2");
        var rects = TestHelper.GetFilledRects(pdfBytes, 1, (0d, 0d, 0d));
        Assert.True(IsFilled(rects, bounds.X + 0.1, bounds.Bottom));
        Assert.True(IsFilled(rects, bounds.Right - 0.1, bounds.Bottom));
        Assert.False(IsFilled(rects, bounds.X - 0.1, bounds.Bottom));
        Assert.False(IsFilled(rects, bounds.Right + 0.1, bounds.Bottom));
    }

    private static bool IsFilled(IReadOnlyList<ReportRect> rects, double x, double y) =>
        rects.Any(r => (r.X <= x) && (x <= r.Right) && (r.Y <= y) && (y <= r.Bottom));
}
