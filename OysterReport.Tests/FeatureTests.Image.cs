namespace OysterReport.Tests;

using ClosedXML.Excel.Drawings;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>画像埋め込みに関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    private static readonly byte[] OnePxPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");

    [Fact]
    public void ImageShouldEmbedSingleImage()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "WithImage";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Logo")
                .MoveTo(sheet.Cell("B2"))
                .WithSize(60, 40);
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldEmbedSingleImage), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("WithImage", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
        Assert.True(pdfBytes.Length > 1000);
    }

    [Fact]
    public void ImageShouldEmbedMultipleImages()
    {
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

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldEmbedMultipleImages), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.True(pdfBytes.Length > 1000);
    }

    [Fact]
    public void ImageShouldHandleFreeFloatingImage()
    {
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

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(ImageShouldHandleFreeFloatingImage), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("FreeFloat", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
