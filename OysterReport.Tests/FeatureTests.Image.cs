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
        // OnePxPng is a grayscale PNG that PDFsharp cannot import directly;
        // it must be embedded through the SkiaSharp transcode fallback
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
        var engine = new OysterReportEngine();
        engine.RenderingOptions.OnRenderWarning = warnings.Add;

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
        var engine = new OysterReportEngine();
        engine.RenderingOptions.OnRenderWarning = warnings.Add;

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
        var engine = new OysterReportEngine();
        engine.RenderingOptions.OnRenderWarning = warnings.Add;

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
}
