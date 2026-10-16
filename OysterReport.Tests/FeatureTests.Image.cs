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
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes));
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
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes));
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
        Assert.Equal(3, TestHelper.GetImageCount(pdfBytes));
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
        Assert.Equal(1, TestHelper.GetImageCount(pdfBytes));
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

    private static byte[] CreatePng(SKColor color)
    {
        using var bitmap = new SKBitmap(2, 2);
        bitmap.Erase(color);
        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }
}
