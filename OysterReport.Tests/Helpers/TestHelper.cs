namespace OysterReport.Tests.Helpers;

using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

internal static class TestHelper
{
    private static readonly string TestOutputDirectory =
        Path.Combine(AppContext.BaseDirectory, "TestOutput");

    public static string IpaExGothicFontPath =>
        Path.Combine(AppContext.BaseDirectory, "ipaexg.ttf");

    public static byte[] GeneratePdfAndSave(string testName, TemplateWorkbook workbook, IReportFontResolver? fontResolver = null)
    {
        var engine = new OysterReportEngine { FontResolver = fontResolver };
        using var ms = new MemoryStream();
        engine.GeneratePdf(workbook, ms);
        var bytes = ms.ToArray();
        SavePdf(testName, bytes);
        return bytes;
    }

    public static byte[] GeneratePdfAndSave(string testName, MemoryStream workbookStream, IReportFontResolver? fontResolver = null)
    {
        workbookStream.Position = 0;
        using var workbook = new TemplateWorkbook(workbookStream);
        return GeneratePdfAndSave(testName, workbook, fontResolver);
    }

    public static byte[] GenerateSheetPdfAndSave(string testName, TemplateSheet sheet, IReportFontResolver? fontResolver = null)
    {
        var engine = new OysterReportEngine { FontResolver = fontResolver };
        using var ms = new MemoryStream();
        engine.GeneratePdf(sheet, ms);
        var bytes = ms.ToArray();
        SavePdf(testName, bytes);
        return bytes;
    }

    public static void SavePdf(string testName, byte[] pdfBytes)
    {
        Directory.CreateDirectory(TestOutputDirectory);
        var safeName = MakeSafeFileName(testName);
        var path = Path.Combine(TestOutputDirectory, safeName + ".pdf");
        File.WriteAllBytes(path, pdfBytes);
    }

    public static string ExtractAllText(byte[] pdfBytes)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        var parts = new List<string>();
        for (var i = 1; i <= doc.NumberOfPages; i++)
        {
            var page = doc.GetPage(i);
            var words = page.GetWords(NearestNeighbourWordExtractor.Instance);
            parts.AddRange(words.Select(static w => w.Text));
        }

        return string.Join(" ", parts);
    }

    public static int GetPageCount(byte[] pdfBytes)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        return doc.NumberOfPages;
    }

    public static IReadOnlyList<Letter> GetLetters(byte[] pdfBytes, int pageNumber = 1)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        return doc.GetPage(pageNumber).Letters;
    }

    public static (double Width, double Height) GetPageSize(byte[] pdfBytes, int pageNumber = 1)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        var page = doc.GetPage(pageNumber);
        return (page.Width, page.Height);
    }

    public static int GetImageCount(byte[] pdfBytes, int pageNumber = 1)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        return doc.GetPage(pageNumber).GetImages().Count();
    }

    public static bool IsValidPdf(byte[] pdfBytes)
    {
        if (pdfBytes.Length < 4)
        {
            return false;
        }

        // %PDF
        return pdfBytes[0] == 0x25 && pdfBytes[1] == 0x50 && pdfBytes[2] == 0x44 && pdfBytes[3] == 0x46;
    }

    private static string MakeSafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return new string(name.Select(c => Array.IndexOf(invalid, c) >= 0 ? '_' : c).ToArray());
    }
}
