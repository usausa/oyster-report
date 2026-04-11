namespace OysterReport.Tests;

using System.Linq;

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

using OysterReport.Internal;
using OysterReport.Tests.Helpers;

using Xunit;

public sealed class ExcelReaderTests
{
    [Fact]
    public void ReadShouldPopulateWorkbookModel()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Title";
            sheet.Cell("A2").Value = "{{Name}}";
            sheet.Range("A1:B1").Merge();
            sheet.Row(2).Hide();
            sheet.PageSetup.CenterHorizontally = true;
            sheet.PageSetup.Header.Left.AddText("Header", XLHFOccurrence.OddPages);
            sheet.PageSetup.Footer.Left.AddText("Footer", XLHFOccurrence.OddPages);
        });

        // Act
        var workbook = ExcelReader.Read(stream);

        // Assert
        var sheet = Assert.Single(workbook.Sheets);
        Assert.Equal("Report", sheet.Name);
        Assert.Single(sheet.MergedRanges);
        Assert.Equal("&LHeader", sheet.HeaderFooter.OddHeader);
        Assert.Equal("&LFooter", sheet.HeaderFooter.OddFooter);
        Assert.True(sheet.PageSetup.CenterHorizontally);
        Assert.Contains(sheet.Rows, row => row.Index == 2 && row.IsHidden);
    }

    [Fact]
    public void ReadShouldPreserveGeneralAlignmentForNumericCells()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Alignment");
            sheet.Cell("A1").Value = 123;
        });

        // Act
        var workbook = ExcelReader.Read(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var cell = sheet.Cells.Single(item => item.Address == "A1");

        // Assert
        Assert.Equal(XLAlignmentHorizontalValues.General, cell.Style.Alignment.Horizontal);
    }

    [Fact]
    public void ReadShouldResolveThemeColorsWithoutThrowing()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Theme");
            var cell = sheet.Cell("A1");
            cell.Value = "Theme";
            cell.Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.4);
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.25);
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.LeftBorderColor = XLColor.FromTheme(XLThemeColor.Accent3, 0.2);
        });

        // Act
        var workbook = ExcelReader.Read(stream);

        // Assert
        var sheet = Assert.Single(workbook.Sheets);
        var cell = sheet.Cells.Single(item => item.Address == "A1");
        Assert.NotEqual("#00000000", cell.Style.Font.ColorHex);
        Assert.NotEqual("#00000000", cell.Style.Fill.BackgroundColorHex);
        Assert.NotEqual("#00000000", cell.Style.Borders.Left.ColorHex);
    }

    [Fact]
    public void ReadShouldSupportFreeFloatingPicturesWithoutBottomRightCell()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Pictures");
            sheet.Cell("A1").Value = "Picture";

            var imageBytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");
            using var imageStream = new MemoryStream(imageBytes, writable: false);
            sheet.AddPicture(imageStream, XLPictureFormat.Png, "Logo")
                .MoveTo(12, 18)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .WithSize(24, 12);
        });

        // Act
        var workbook = ExcelReader.Read(stream);

        // Assert
        var sheet = Assert.Single(workbook.Sheets);
        var image = Assert.Single(sheet.Images);
        Assert.Equal("Logo", image.Name);
        Assert.Null(image.ToCellAddress);
    }

    [Fact]
    public void ReadShouldConvertPageMarginsToPoints()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Margins");
            sheet.Cell("A1").Value = "Margin";
            sheet.PageSetup.Margins.Left = 1.25;
            sheet.PageSetup.Margins.Top = 0.5;
            sheet.PageSetup.Margins.Right = 0.75;
            sheet.PageSetup.Margins.Bottom = 1.0;
            sheet.PageSetup.Margins.Header = 0.3;
            sheet.PageSetup.Margins.Footer = 0.2;
        });

        // Act
        var workbook = ExcelReader.Read(stream);

        // Assert
        var sheet = Assert.Single(workbook.Sheets);

        Assert.Equal(90d, sheet.PageSetup.Margins.Left, 3);
        Assert.Equal(36d, sheet.PageSetup.Margins.Top, 3);
        Assert.Equal(54d, sheet.PageSetup.Margins.Right, 3);
        Assert.Equal(72d, sheet.PageSetup.Margins.Bottom, 3);
        Assert.Equal(21.6d, sheet.PageSetup.HeaderMarginPoint, 3);
        Assert.Equal(14.4d, sheet.PageSetup.FooterMarginPoint, 3);
    }

    [Fact]
    public void ReadShouldApplyTableStripeFillFromExcelTableTheme()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Table");
            sheet.Cell("A1").Value = "No.";
            sheet.Cell("B1").Value = "Name";
            sheet.Cell("A2").Value = 1;
            sheet.Cell("B2").Value = "First";
            sheet.Cell("A3").Value = 2;
            sheet.Cell("B3").Value = "Second";

            var table = sheet.Range("A1:B3").CreateTable();
            table.Theme = XLTableTheme.TableStyleLight4;
            table.ShowRowStripes = true;
        });

        // Act
        var workbook = ExcelReader.Read(stream);

        // Assert
        var sheet = Assert.Single(workbook.Sheets);
        var firstDataRowCell = sheet.Cells.Single(cell => cell.Address == "A2");
        var secondDataRowCell = sheet.Cells.Single(cell => cell.Address == "A3");

        Assert.NotEqual("#00000000", firstDataRowCell.Style.Fill.BackgroundColorHex);
        Assert.NotEqual(firstDataRowCell.Style.Fill.BackgroundColorHex, secondDataRowCell.Style.Fill.BackgroundColorHex);
    }

    [Fact]
    public void ReadShouldApplyDifferentStripeColorsForDifferentTableStyles()
    {
        // Arrange: two tables on the same sheet using different accent-based styles
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Tables");

            sheet.Cell("A1").Value = "No.";
            sheet.Cell("A2").Value = 1;
            sheet.Cell("A3").Value = 2;
            var light4 = sheet.Range("A1:A3").CreateTable();
            light4.Theme = XLTableTheme.TableStyleLight4;
            light4.ShowRowStripes = true;

            sheet.Cell("C1").Value = "No.";
            sheet.Cell("C2").Value = 1;
            sheet.Cell("C3").Value = 2;
            var light6 = sheet.Range("C1:C3").CreateTable();
            light6.Theme = XLTableTheme.TableStyleLight6;
            light6.ShowRowStripes = true;
        });

        // Act
        var workbook = ExcelReader.Read(stream);
        var sheet = Assert.Single(workbook.Sheets);

        var light4Cell = sheet.Cells.Single(cell => cell.Address == "A2");
        var light6Cell = sheet.Cells.Single(cell => cell.Address == "C2");

        // Assert: both styles apply a stripe fill, and the two fills are different
        Assert.NotEqual("#00000000", light4Cell.Style.Fill.BackgroundColorHex);
        Assert.NotEqual("#00000000", light6Cell.Style.Fill.BackgroundColorHex);
        Assert.NotEqual(light4Cell.Style.Fill.BackgroundColorHex, light6Cell.Style.Fill.BackgroundColorHex);
    }

    [Fact]
    public void ReadShouldSkipStripeForTableStyleWithoutCatalogEntry()
    {
        // Arrange: use a neutral style (Light1) that has no accent stripe fill
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Table");
            sheet.Cell("A1").Value = "No.";
            sheet.Cell("A2").Value = 1;
            var table = sheet.Range("A1:A2").CreateTable();
            table.Theme = XLTableTheme.TableStyleLight1;
            table.ShowRowStripes = true;
        });

        // Act
        var workbook = ExcelReader.Read(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var dataCell = sheet.Cells.Single(cell => cell.Address == "A2");

        // Assert: no stripe fill is applied because Light1 is not in the catalog
        Assert.Equal("#00000000", dataCell.Style.Fill.BackgroundColorHex);
    }
}
