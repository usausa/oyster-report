// <copyright file="PdfTestHelper.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests.Helpers;

using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

/// <summary>
/// テスト用PDFヘルパー: PDF生成・保存・テキスト抽出を提供する。
/// </summary>
internal static class PdfTestHelper
{
    private static readonly string TestOutputDirectory =
        Path.Combine(AppContext.BaseDirectory, "TestOutput");

    public static string IpaExGothicFontPath =>
        Path.Combine(AppContext.BaseDirectory, "ipaexg.ttf");

    /// <summary>
    /// TemplateWorkbook から PDF バイト列を生成し、テスト名を元にしたファイル名で保存する。
    /// </summary>
    public static byte[] GeneratePdfAndSave(string testName, TemplateWorkbook workbook, IReportFontResolver? fontResolver = null)
    {
        var engine = new OysterReportEngine { FontResolver = fontResolver };
        using var ms = new MemoryStream();
        engine.GeneratePdf(workbook, ms);
        var bytes = ms.ToArray();
        SavePdf(testName, bytes);
        return bytes;
    }

    /// <summary>
    /// MemoryStream からテンプレートを開いて PDF を生成・保存する。
    /// </summary>
    public static byte[] GeneratePdfAndSave(string testName, MemoryStream workbookStream, IReportFontResolver? fontResolver = null)
    {
        workbookStream.Position = 0;
        using var workbook = new TemplateWorkbook(workbookStream);
        return GeneratePdfAndSave(testName, workbook, fontResolver);
    }

    /// <summary>
    /// TemplateSheet から PDF バイト列を生成し、テスト名を元にしたファイル名で保存する。
    /// </summary>
    public static byte[] GenerateSheetPdfAndSave(string testName, TemplateSheet sheet, IReportFontResolver? fontResolver = null)
    {
        var engine = new OysterReportEngine { FontResolver = fontResolver };
        using var ms = new MemoryStream();
        engine.GeneratePdf(sheet, ms);
        var bytes = ms.ToArray();
        SavePdf(testName, bytes);
        return bytes;
    }

    /// <summary>
    /// PDF バイト列をテスト名ファイルとして保存する。
    /// </summary>
    public static void SavePdf(string testName, byte[] pdfBytes)
    {
        Directory.CreateDirectory(TestOutputDirectory);
        var safeName = MakeSafeFileName(testName);
        var path = Path.Combine(TestOutputDirectory, safeName + ".pdf");
        File.WriteAllBytes(path, pdfBytes);
    }

    /// <summary>
    /// PDF バイト列から全ページのテキストを取得する。
    /// </summary>
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

    /// <summary>
    /// PDF バイト列から指定ページ番号 (1-based) のテキストを取得する。
    /// </summary>
    public static string ExtractPageText(byte[] pdfBytes, int pageNumber)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        var page = doc.GetPage(pageNumber);
        var words = page.GetWords(NearestNeighbourWordExtractor.Instance);
        return string.Join(" ", words.Select(static w => w.Text));
    }

    /// <summary>
    /// PDF バイト列からページ数を取得する。
    /// </summary>
    public static int GetPageCount(byte[] pdfBytes)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        return doc.NumberOfPages;
    }

    /// <summary>
    /// PDF バイト列から指定ページの Letter リストを取得する。
    /// </summary>
    public static IReadOnlyList<Letter> GetLetters(byte[] pdfBytes, int pageNumber = 1)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        return doc.GetPage(pageNumber).Letters;
    }

    /// <summary>
    /// PDF バイト列から指定ページの (幅, 高さ) をポイント単位で取得する。
    /// </summary>
    public static (double Width, double Height) GetPageSize(byte[] pdfBytes, int pageNumber = 1)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        var page = doc.GetPage(pageNumber);
        return (page.Width, page.Height);
    }

    /// <summary>
    /// PDF バイト列から指定ページの画像一覧を取得する。
    /// </summary>
    public static int GetImageCount(byte[] pdfBytes, int pageNumber = 1)
    {
        using var doc = PdfDocument.Open(pdfBytes);
        return doc.GetPage(pageNumber).GetImages().Count();
    }

    /// <summary>
    /// PDF が %PDF ヘッダーで始まり空でないことを確認する。
    /// </summary>
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
