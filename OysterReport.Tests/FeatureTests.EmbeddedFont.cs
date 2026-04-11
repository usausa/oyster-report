// <copyright file="FeatureTests.EmbeddedFont.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>IPAex ゴシック埋め込みフォントに関する機能テスト。</summary>
public sealed partial class FeatureTests
{
    [Fact]
    public void EmbeddedFontShouldEmbedIpaExGothicForJapanese()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "日本語テスト";
            cell.Style.Font.FontName = "ＭＳ Ｐゴシック";
            cell.Style.Font.FontSize = 12d;
        });

        var resolver = new IpaExGothicFontResolver(PdfTestHelper.IpaExGothicFontPath);
        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(EmbeddedFontShouldEmbedIpaExGothicForJapanese),
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var letters = PdfTestHelper.GetLetters(pdfBytes);
        Assert.Contains(
            letters,
            static l => l.FontName is not null && l.FontName.Contains("IPAex", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void EmbeddedFontShouldRenderMultipleJapaneseCells()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            foreach (var (row, text) in new[]
            {
                (1, "請求書"),
                (2, "合計金額"),
                (3, "お支払い期限")
            })
            {
                var cell = sheet.Cell(row, 1);
                cell.Value = text;
                cell.Style.Font.FontName = "ＭＳ Ｐゴシック";
                cell.Style.Font.FontSize = 11d;
            }
        });

        var resolver = new IpaExGothicFontResolver(PdfTestHelper.IpaExGothicFontPath);
        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(EmbeddedFontShouldRenderMultipleJapaneseCells),
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void EmbeddedFontShouldRenderJapaneseBold()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "太字テスト";
            cell.Style.Font.FontName = "ＭＳ Ｐゴシック";
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 14d;
        });

        var resolver = new IpaExGothicFontResolver(PdfTestHelper.IpaExGothicFontPath);
        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(EmbeddedFontShouldRenderJapaneseBold),
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void EmbeddedFontShouldRenderMixedJapaneseAndAscii()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell1 = sheet.Cell("A1");
            cell1.Value = "合計: 12345円";
            cell1.Style.Font.FontName = "ＭＳ Ｐゴシック";
            cell1.Style.Font.FontSize = 11d;
            var cell2 = sheet.Cell("A2");
            cell2.Value = "EnglishText";
            cell2.Style.Font.FontSize = 11d;
        });

        var resolver = new IpaExGothicFontResolver(PdfTestHelper.IpaExGothicFontPath);
        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(EmbeddedFontShouldRenderMixedJapaneseAndAscii),
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("EnglishText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void EmbeddedFontShouldLoadFontFromPath()
    {
        Assert.True(
            File.Exists(PdfTestHelper.IpaExGothicFontPath),
            $"ipaexg.ttf not found at: {PdfTestHelper.IpaExGothicFontPath}");

        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "フォントファイル確認";
            cell.Style.Font.FontName = "メイリオ";
        });

        var resolver = new IpaExGothicFontResolver(PdfTestHelper.IpaExGothicFontPath);
        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            nameof(EmbeddedFontShouldLoadFontFromPath),
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}
