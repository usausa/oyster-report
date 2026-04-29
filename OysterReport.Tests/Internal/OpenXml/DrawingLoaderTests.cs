namespace OysterReport.Tests.Internal.OpenXml;

using DocumentFormat.OpenXml.Packaging;

using OysterReport.Internal;

public sealed class DrawingLoaderTests
{
    private static readonly byte[] OnePxPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");

    //--------------------------------------------------------------------------------
    // No drawings
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldReturnEmptyWhenWorksheetHasNoDrawingsPart()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
        });

        // Act
        var images = LoadDrawings(stream);

        // Assert
        Assert.Empty(images);
    }

    //--------------------------------------------------------------------------------
    // twoCellAnchor (default placement)
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldExtractSingleTwoCellAnchorImage()
    {
        // Arrange — placing a picture across a from/to cell range yields a twoCellAnchor
        // without an editAs override, so the loader retains the ToCellAddress.
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Anchor";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Logo")
                .MoveTo(sheet.Cell("B2"), sheet.Cell("D5"));
        });

        // Act
        var images = LoadDrawings(stream);
        var image = Assert.Single(images);

        // Assert
        Assert.Equal("Logo", image.Name);
        Assert.Equal("B2", image.FromCellAddress);
        Assert.Equal("D5", image.ToCellAddress);
        Assert.True(image.WidthPoint > 0);
        Assert.True(image.HeightPoint > 0);
        Assert.False(image.ImageBytes.IsEmpty);
        Assert.Equal(OnePxPng.Length, image.ImageBytes.Length);
        Assert.Equal(OnePxPng, image.ImageBytes.ToArray());
    }

    [Fact]
    public void LoadShouldExtractMultipleImages()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            using var img1 = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(img1, XLPictureFormat.Png, "First")
                .MoveTo(sheet.Cell("A1"))
                .WithSize(40, 30);
            using var img2 = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(img2, XLPictureFormat.Png, "Second")
                .MoveTo(sheet.Cell("D1"))
                .WithSize(40, 30);
        });

        // Act
        var images = LoadDrawings(stream);

        // Assert
        Assert.Equal(2, images.Length);
        Assert.Contains(images, i => i is { Name: "First", FromCellAddress: "A1" });
        Assert.Contains(images, i => i is { Name: "Second", FromCellAddress: "D1" });
    }

    //--------------------------------------------------------------------------------
    // oneCellAnchor (Move placement keeps the image anchored without resize)
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldExtractOneCellAnchorImage()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Pinned")
                .MoveTo(sheet.Cell("C3"))
                .WithPlacement(XLPicturePlacement.Move)
                .WithSize(50, 30);
        });

        // Act
        var images = LoadDrawings(stream);
        var image = Assert.Single(images);

        // Assert — oneCellAnchor leaves ToCellAddress null because the image only has a single anchor.
        Assert.Equal("Pinned", image.Name);
        Assert.Equal("C3", image.FromCellAddress);
        Assert.Null(image.ToCellAddress);
        Assert.True(image.WidthPoint > 0);
        Assert.True(image.HeightPoint > 0);
    }

    //--------------------------------------------------------------------------------
    // absoluteAnchor (FreeFloating) — DrawingLoader does not surface these.
    //--------------------------------------------------------------------------------

    [Fact]
    public void LoadShouldNotExtractAbsoluteAnchorFreeFloatingImage()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Floater")
                .MoveTo(20, 30)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .WithSize(50, 30);
        });

        // Act
        var images = LoadDrawings(stream);

        // Assert
        Assert.Empty(images);
    }

    //--------------------------------------------------------------------------------
    // Helpers
    //--------------------------------------------------------------------------------

    private static ReportImage[] LoadDrawings(Stream stream)
    {
        stream.Position = 0;
        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        var wsPart = doc.WorkbookPart!.WorksheetParts.First();
        return DrawingLoader.Load(wsPart).ToArray();
    }
}
