namespace OysterReport.Tests;

using SkiaSharp;

public sealed partial class FeatureTests
{
    private static readonly byte[] OnePxPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");

    [Fact]
    public void ImageShouldEmbedSingleImage()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "WithImage";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Logo")
                .MoveTo(sheet.Cell("B2"))
                .WithSize(60, 40);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldEmbedSingleImage), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("WithImage", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
        Assert.True(pdfBytes.Length > 1000);
    }

    [Fact]
    public void ImageShouldEmbedMultipleImages()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "MultiImage";
            using var img1 = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(img1, XLPictureFormat.Png, "Image1")
                .MoveTo(sheet.Cell("B1"))
                .WithSize(40, 30);
            using var img2 = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(img2, XLPictureFormat.Png, "Image2")
                .MoveTo(sheet.Cell("D1"))
                .WithSize(40, 30);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldEmbedMultipleImages), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.True(pdfBytes.Length > 1000);
    }

    [Fact]
    public void ImageShouldEmbedPngInsideUsedRange()
    {
        // OnePxPng is a grayscale PNG; it is embedded as-is (DeviceGray) without decoding
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "PngInRange";
            sheet.Cell("C5").Value = "End";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Png")
                .MoveTo(sheet.Cell("B2"))
                .WithSize(40, 30);
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var warnings = new List<ReportRenderWarning>();
        var engine = new OysterReportEngine { RenderingOptions = { OnRenderWarning = warnings.Add } };

        // Act
        using var output = new MemoryStream();
        engine.GeneratePdf(workbook, output);
        var pdfBytes = output.ToArray();
        TestHelper.SavePdf(nameof(ImageShouldEmbedPngInsideUsedRange), pdfBytes);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.True(
            warnings.Count == 0,
            "warnings: " + String.Join("; ", warnings.Select(static w => $"{w.Source} {w.Exception?.GetType().Name}: {w.Exception?.Message}")));
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes, 1));
    }

    [Fact]
    public void ImageShouldEmbedJpegInsideUsedRange()
    {
        // JPEG is imported by PDFsharp directly; also guards the XImage disposal lifecycle
        byte[] jpegBytes;
        using (var bitmap = new SKBitmap(2, 2))
        {
            bitmap.Erase(SKColors.Red);
            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Jpeg, 90);
            jpegBytes = data.ToArray();
        }

        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "JpegInRange";
            sheet.Cell("C5").Value = "End";
            using var imgStream = new MemoryStream(jpegBytes, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Jpeg, "Jpeg")
                .MoveTo(sheet.Cell("B2"))
                .WithSize(40, 30);
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var warnings = new List<ReportRenderWarning>();
        var engine = new OysterReportEngine { RenderingOptions = { OnRenderWarning = warnings.Add } };

        // Act
        using var output = new MemoryStream();
        engine.GeneratePdf(workbook, output);
        var pdfBytes = output.ToArray();
        TestHelper.SavePdf(nameof(ImageShouldEmbedJpegInsideUsedRange), pdfBytes);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Empty(warnings);
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes, 1));
    }

    [Fact]
    public void ImageShouldNotifyWarningWhenImageDecodeFails()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BrokenImage";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        workbook.ReportWorkbook.Sheets[0].AddImage(new ReportImage
        {
            Name = "Broken",
            FromCellAddress = "A1",
            WidthPoint = 40d,
            HeightPoint = 30d,
            ImageBytes = new byte[] { 0x01, 0x02, 0x03 }
        });

        var warnings = new List<ReportRenderWarning>();
        var engine = new OysterReportEngine { RenderingOptions = { OnRenderWarning = warnings.Add } };

        // Act
        using var output = new MemoryStream();
        engine.GeneratePdf(workbook, output);
        var pdfBytes = output.ToArray();
        TestHelper.SavePdf(nameof(ImageShouldNotifyWarningWhenImageDecodeFails), pdfBytes);

        // Assert — the broken image is skipped, the PDF stays valid, and the warning reports the source
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BrokenImage", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
        var warning = Assert.Single(warnings);
        Assert.Equal(ReportRenderWarningKind.ImageDecodeFailed, warning.Kind);
        Assert.Equal("Report", warning.SheetName);
        Assert.Equal("Broken", warning.Source);
        Assert.Contains("Broken", warning.Message, StringComparison.Ordinal);
        Assert.NotNull(warning.Exception);
    }

    [Fact]
    public void ImageShouldHandleFreeFloatingImage()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "FreeFloat";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "FreeImg")
                .MoveTo(20, 30)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .WithSize(50, 30);
        });

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldHandleFreeFloatingImage), stream);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("FreeFloat", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void ImageShouldEmbedIdenticalImagesOnce()
    {
        // Arrange
        var single = GenerateImagePdf((OnePxPng, "Report", "B2"));

        // Act
        var pdfBytes = GenerateImagePdf((OnePxPng, "Report", "B2"), (OnePxPng, "Report", "D2"), (OnePxPng, "Report", "B5"));
        TestHelper.SavePdf(nameof(ImageShouldEmbedIdenticalImagesOnce), pdfBytes);

        // Assert — one XObject set, drawn three times
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(CountImageXObjects(single), CountImageXObjects(pdfBytes));
        Assert.Equal(3, TestHelper.GetImageCount(pdfBytes, 1));
    }

    [Fact]
    public void ImageShouldEmbedIdenticalImagesOnceAcrossSheets()
    {
        // Arrange
        var single = GenerateImagePdf((OnePxPng, "Sheet1", "B2"));

        // Act
        var pdfBytes = GenerateImagePdf((OnePxPng, "Sheet1", "B2"), (OnePxPng, "Sheet2", "B2"));
        TestHelper.SavePdf(nameof(ImageShouldEmbedIdenticalImagesOnceAcrossSheets), pdfBytes);

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(CountImageXObjects(single), CountImageXObjects(pdfBytes));
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes, 1));
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes, 2));
    }

    [Fact]
    public void ImageShouldEmbedDistinctImagesSeparately()
    {
        // Arrange
        var other = CreatePng(SKColors.Red);
        var first = GenerateImagePdf((OnePxPng, "Report", "B2"));
        var second = GenerateImagePdf((other, "Report", "B2"));

        // Act
        var pdfBytes = GenerateImagePdf((OnePxPng, "Report", "B2"), (other, "Report", "D2"));

        // Assert
        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.True(CountImageXObjects(first) > 0);
        Assert.Equal(CountImageXObjects(first) + CountImageXObjects(second), CountImageXObjects(pdfBytes));
    }

    [Fact]
    public void ImageShouldNotifyWarningForEachPlacementOfBrokenImage()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Broken";
            sheet.Cell("C3").Value = "End";
        });
        using var workbook = new TemplateWorkbook(stream);
        byte[] broken = [0x01, 0x02, 0x03];
        foreach (var (name, address) in new[] { ("First", "A1"), ("Second", "B2") })
        {
            workbook.ReportWorkbook.Sheets[0].AddImage(new ReportImage
            {
                Name = name,
                FromCellAddress = address,
                WidthPoint = 40d,
                HeightPoint = 30d,
                ImageBytes = broken
            });
        }

        var warnings = new List<ReportRenderWarning>();
        var engine = new OysterReportEngine { RenderingOptions = { OnRenderWarning = warnings.Add } };

        // Act
        using var output = new MemoryStream();
        engine.GeneratePdf(workbook, output);

        // Assert — failures are not cached, so every placement is reported
        Assert.True(TestHelper.IsValidPdf(output.ToArray()));
        Assert.Equal(["First", "Second"], warnings.Select(static x => x.Source));
    }

    //--------------------------------------------------------------------------------
    // Insertion
    //--------------------------------------------------------------------------------

    [Fact]
    public void ImageShouldReplacePlaceholderAndFitMergedCells()
    {
        // Arrange
        using var stream = CreatePlaceholderImageWorkbook();
        using var workbook = new TemplateWorkbook(stream);
        var sheet = workbook.GetSheet("Report");

        // Act
        var count = sheet.ReplacePlaceholderWithImage("QR", CreatePng(SKColors.Black));
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldReplacePlaceholderAndFitMergedCells), workbook);

        // Assert: the square image keeps its aspect ratio in the center of B2:C3
        Assert.Equal(1, count);
        Assert.Equal(string.Empty, sheet.GetCellText(2, 2));
        var cell = GetCellBounds(workbook, "B2");
        var image = Assert.Single(TestHelper.GetImageBounds(pdfBytes, 1));
        var size = Math.Min(cell.Width, cell.Height);
        var pageHeight = TestHelper.GetPageSize(pdfBytes, 1).Height;
        Assert.Equal(size, image.Width, 2);
        Assert.Equal(size, image.Height, 2);
        Assert.Equal(cell.X + (cell.Width / 2), image.Left + (image.Width / 2), 2);
        Assert.Equal(pageHeight - (cell.Y + (cell.Height / 2)), image.Bottom + (image.Height / 2), 2);
    }

    [Fact]
    public void ImageShouldFillCellWhenFitIsFill()
    {
        // Arrange
        using var stream = CreatePlaceholderImageWorkbook();
        using var workbook = new TemplateWorkbook(stream);

        // Act
        workbook.GetSheet("Report").SetCellImage(1, 1, CreatePng(SKColors.Red), ReportImageFit.Fill);
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldFillCellWhenFitIsFill), workbook);

        // Assert
        var cell = GetCellBounds(workbook, "A1");
        var image = Assert.Single(TestHelper.GetImageBounds(pdfBytes, 1));
        Assert.Equal(cell.Width, image.Width, 2);
        Assert.Equal(cell.Height, image.Height, 2);
    }

    [Fact]
    public void ImageInCellShouldBeDrawnUnderBorders()
    {
        // Arrange
        using var stream = CreatePlaceholderImageWorkbook();
        using var workbook = new TemplateWorkbook(stream);
        workbook.ReplacePlaceholderWithImage("QR", CreatePng(SKColors.Black));

        // Act
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageInCellShouldBeDrawnUnderBorders), workbook);

        // Assert: the borders of B2:C3 are filled rectangles drawn after the image
        var operators = TestHelper.GetOperators(pdfBytes, 1);
        var imageIndex = operators.ToList().IndexOf("Do");
        Assert.True(imageIndex >= 0);
        Assert.Contains("f", operators.Skip(imageIndex + 1));
    }

    [Fact]
    public void ImageShouldReplacePlaceholderInEachCopiedRow()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Item";
            sheet.Cell("B1").Value = "{{Photo}}";
            sheet.Cell("C10").Value = "End";
        });
        using var workbook = new TemplateWorkbook(stream);
        var templateRow = workbook.GetSheet("Report").GetRow(1);

        // Act
        var last = templateRow;
        foreach (var color in new[] { SKColors.Red, SKColors.Green, SKColors.Blue })
        {
            var row = templateRow.InsertCopyAfter(last);
            row.ReplacePlaceholderWithImage("Photo", CreatePng(color));
            last = row;
        }

        templateRow.Delete();
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldReplacePlaceholderInEachCopiedRow), workbook);

        // Assert
        Assert.Equal(3, TestHelper.GetImageCount(pdfBytes, 1));
        Assert.DoesNotContain("{{Photo}}", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void ImageShouldReplacePlaceholderOnEverySheet()
    {
        // Arrange
        using var stream = CreatePlaceholderImageWorkbook();
        using var workbook = new TemplateWorkbook(stream);
        workbook.CopySheet("Report", "Copy");

        // Act
        var count = workbook.ReplacePlaceholderWithImage("QR", CreatePng(SKColors.Black));
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldReplacePlaceholderOnEverySheet), workbook);

        // Assert
        Assert.Equal(2, count);
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes, 1));
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes, 2));
    }

    [Fact]
    public void ImageShouldKeepCellPlacementWhenSheetIsCopied()
    {
        // Arrange
        using var stream = CreatePlaceholderImageWorkbook();
        using var workbook = new TemplateWorkbook(stream);
        workbook.GetSheet("Report").ReplacePlaceholderWithImage("QR", CreatePng(SKColors.Black), ReportImageFit.Fill);

        // Act
        var copy = workbook.CopySheet("Report", "Copy");

        // Assert
        var image = Assert.Single(copy.UnderlyingSheet.Images);
        Assert.True(image.InCell);
        Assert.Equal(ReportImageFit.Fill, image.Fit);
        Assert.Equal("B2", image.FromCellAddress);
        Assert.Equal("C3", image.ToCellAddress);
    }

    [Fact]
    public void ImageShouldShrinkWhenRowsInsideItsCellsAreDeleted()
    {
        // Arrange
        using var stream = TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("B2").Value = "{{QR}}";
            sheet.Range("B2:C5").Merge();
            sheet.Cell("E8").Value = "End";
        });
        using var workbook = new TemplateWorkbook(stream);
        var sheet = workbook.GetSheet("Report");
        sheet.ReplacePlaceholderWithImage("QR", CreatePng(SKColors.Black));

        // Act
        sheet.DeleteRows(3, 4);

        // Assert
        var image = Assert.Single(sheet.UnderlyingSheet.Images);
        Assert.Equal("B2", image.FromCellAddress);
        Assert.Equal("C3", image.ToCellAddress);
    }

    [Fact]
    public void ImageShouldReplaceNamedImageInItsFrame()
    {
        // Arrange
        using var stream = CreateNamedImageWorkbook();
        using var workbook = new TemplateWorkbook(stream);
        var frame = workbook.ReportWorkbook.Sheets[0].Images[0];

        // Act
        var count = workbook.ReplaceImage("Stamp", CreatePng(SKColors.Red, 4));
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldReplaceNamedImageInItsFrame), workbook);

        // Assert: the 2:1 image is scaled to the frame width
        Assert.Equal(1, count);
        var image = Assert.Single(TestHelper.GetImageBounds(pdfBytes, 1));
        Assert.Equal(frame.WidthPoint, image.Width, 2);
        Assert.Equal(frame.WidthPoint / 2, image.Height, 2);
    }

    [Fact]
    public void ImageShouldRemoveNamedImage()
    {
        // Arrange
        using var stream = CreateNamedImageWorkbook();
        using var workbook = new TemplateWorkbook(stream);

        // Act
        var count = workbook.RemoveImage("Stamp");
        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldRemoveNamedImage), workbook);

        // Assert
        Assert.Equal(1, count);
        Assert.Equal(0, TestHelper.GetImageCount(pdfBytes, 1));
    }

    // B2:C3 is merged, holds {{QR}} and has an outside border
    private static MemoryStream CreatePlaceholderImageWorkbook() =>
        TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Start";
            sheet.Cell("B2").Value = "{{QR}}";
            sheet.Range("B2:C3").Merge();
            sheet.Range("B2:C3").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell("E6").Value = "End";
        });

    private static MemoryStream CreateNamedImageWorkbook() =>
        TestWorkbookFactory.CreateWorkbook(static workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Start";
            sheet.Cell("F8").Value = "End";
            using var image = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(image, XLPictureFormat.Png, "Stamp").MoveTo(sheet.Cell("B2")).WithSize(40, 30);
        });

    private static ReportRect GetCellBounds(TemplateWorkbook workbook, string address) =>
        PdfRenderPlanner.BuildPlan(workbook.ReportWorkbook)[0].Pages[0].Cells.Single(x => x.CellAddress == address).OuterBounds;

    private static byte[] GenerateImagePdf(params (byte[] Image, string Sheet, string Cell)[] placements)
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var index = 0;
            foreach (var (image, sheetName, cell) in placements)
            {
                var sheet = workbook.Worksheets.TryGetWorksheet(sheetName, out var existing) ? existing : workbook.AddWorksheet(sheetName);
                sheet.Cell("A1").Value = "Start";
                sheet.Cell("F8").Value = "End";
                using var imageStream = new MemoryStream(image, writable: false);
                sheet.AddPicture(imageStream, XLPictureFormat.Png, $"Image{++index}")
                    .MoveTo(sheet.Cell(cell))
                    .WithSize(40, 30);
            }
        });

        using var workbook = new TemplateWorkbook(stream);
        using var output = new MemoryStream();
        new OysterReportEngine().GeneratePdf(workbook, output);
        return output.ToArray();
    }

    private static int CountImageXObjects(byte[] pdfBytes)
    {
        var text = Encoding.Latin1.GetString(pdfBytes);
        var count = 0;
        for (var index = text.IndexOf("/Subtype /Image", StringComparison.Ordinal); index >= 0; index = text.IndexOf("/Subtype /Image", index + 1, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    private static byte[] CreatePng(SKColor color, int width = 2, int height = 2)
    {
        using var bitmap = new SKBitmap(width, height);
        bitmap.Erase(color);
        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }
}
