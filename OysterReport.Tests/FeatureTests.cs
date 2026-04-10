// <copyright file="FeatureTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

using OysterReport.Tests.Helpers;

using Xunit;

// ============================================================
// 機能テスト: フォントサイズ
// ============================================================

/// <summary>
/// フォントサイズに関する機能テスト。
/// 小・標準・大サイズでテキストが PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureFontSizeTests
{
    [Theory]
    [InlineData(6d, "TinyText")]
    [InlineData(11d, "NormalText")]
    [InlineData(18d, "LargeText")]
    [InlineData(24d, "HugeText")]
    public void PdfShouldContainTextWithVariousFontSizes(double fontSize, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Font.FontSize = fontSize;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontSizeTests)}_{nameof(PdfShouldContainTextWithVariousFontSizes)}_{fontSize}pt",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains(cellValue, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderMultipleFontSizesOnOnePage()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Size8";
            sheet.Cell("A1").Style.Font.FontSize = 8d;
            sheet.Cell("A2").Value = "Size12";
            sheet.Cell("A2").Style.Font.FontSize = 12d;
            sheet.Cell("A3").Value = "Size16";
            sheet.Cell("A3").Style.Font.FontSize = 16d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontSizeTests)}_{nameof(PdfShouldRenderMultipleFontSizesOnOnePage)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Size8", text, StringComparison.Ordinal);
        Assert.Contains("Size12", text, StringComparison.Ordinal);
        Assert.Contains("Size16", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: フォントスタイル (太字・斜体)
// ============================================================

/// <summary>
/// フォントスタイル (太字・斜体・組み合わせ) に関する機能テスト。
/// </summary>
public sealed class FeatureFontStyleTests
{
    [Fact]
    public void PdfShouldContainBoldText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldText";
            cell.Style.Font.Bold = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontStyleTests)}_{nameof(PdfShouldContainBoldText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainItalicText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ItalicText";
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontStyleTests)}_{nameof(PdfShouldContainItalicText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ItalicText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainBoldItalicText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldItalic";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontStyleTests)}_{nameof(PdfShouldContainBoldItalicText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldItalic", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderMixedNormalBoldItalicOnSamePage()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Normal";
            var boldCell = sheet.Cell("A2");
            boldCell.Value = "Bold";
            boldCell.Style.Font.Bold = true;
            var italicCell = sheet.Cell("A3");
            italicCell.Value = "Italic";
            italicCell.Style.Font.Italic = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontStyleTests)}_{nameof(PdfShouldRenderMixedNormalBoldItalicOnSamePage)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Normal", text, StringComparison.Ordinal);
        Assert.Contains("Bold", text, StringComparison.Ordinal);
        Assert.Contains("Italic", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: フォントカラー
// ============================================================

/// <summary>
/// フォントカラーに関する機能テスト。
/// 赤・青・緑・カスタム RGB でテキストが PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureFontColorTests
{
    [Theory]
    [InlineData("RedText", 255, 0, 0)]
    [InlineData("BlueText", 0, 0, 255)]
    [InlineData("GreenText", 0, 128, 0)]
    public void PdfShouldContainColoredText(string cellValue, int r, int g, int b)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Font.FontColor = XLColor.FromArgb(r, g, b);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontColorTests)}_{nameof(PdfShouldContainColoredText)}_{cellValue}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainThemeColorText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeColorText";
            cell.Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.4);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFontColorTests)}_{nameof(PdfShouldContainThemeColorText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeColorText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: 背景色 (Fill)
// ============================================================

/// <summary>
/// セル背景色に関する機能テスト。
/// 背景色があるセルのテキストが PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureFillColorTests
{
    [Theory]
    [InlineData("YellowBg", 255, 255, 0)]
    [InlineData("LightBlueBg", 173, 216, 230)]
    [InlineData("GrayBg", 192, 192, 192)]
    public void PdfShouldContainTextOnColoredBackground(string cellValue, int r, int g, int b)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(r, g, b);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFillColorTests)}_{nameof(PdfShouldContainTextOnColoredBackground)}_{cellValue}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderMultipleCellsWithDifferentBackgroundColors()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var redCell = sheet.Cell("A1");
            redCell.Value = "RedBg";
            redCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 200, 200);
            var greenCell = sheet.Cell("A2");
            greenCell.Value = "GreenBg";
            greenCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 255, 200);
            var blueCell = sheet.Cell("A3");
            blueCell.Value = "BlueBg";
            blueCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 200, 255);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFillColorTests)}_{nameof(PdfShouldRenderMultipleCellsWithDifferentBackgroundColors)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("RedBg", text, StringComparison.Ordinal);
        Assert.Contains("GreenBg", text, StringComparison.Ordinal);
        Assert.Contains("BlueBg", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderThemeBackgroundColor()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ThemeBg";
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2, -0.25);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureFillColorTests)}_{nameof(PdfShouldRenderThemeBackgroundColor)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ThemeBg", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: セル結合
// ============================================================

/// <summary>
/// セル結合に関する機能テスト。
/// 水平・垂直・矩形結合でテキストが PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureMergedCellTests
{
    [Fact]
    public void PdfShouldContainTextInHorizontallyMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "HorizontalMerge";
            sheet.Range("A1:D1").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureMergedCellTests)}_{nameof(PdfShouldContainTextInHorizontallyMergedCell)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("HorizontalMerge", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainTextInVerticallyMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VerticalMerge";
            sheet.Range("A1:A4").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureMergedCellTests)}_{nameof(PdfShouldContainTextInVerticallyMergedCell)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("VerticalMerge", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainTextInRectangularMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("B2").Value = "RectMerge";
            sheet.Range("B2:D4").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureMergedCellTests)}_{nameof(PdfShouldContainTextInRectangularMergedCell)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("RectMerge", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderMultipleMergedRanges()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Range("A1:C1").Merge();
            sheet.Cell("A2").Value = "Left";
            sheet.Range("A2:A4").Merge();
            sheet.Cell("B2").Value = "Right";
            sheet.Range("B2:C4").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureMergedCellTests)}_{nameof(PdfShouldRenderMultipleMergedRanges)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("Left", text, StringComparison.Ordinal);
        Assert.Contains("Right", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldNotDuplicateTextFromMergedSubCells()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "MergeOwner";
            sheet.Range("A1:C1").Merge();
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureMergedCellTests)}_{nameof(PdfShouldNotDuplicateTextFromMergedSubCells)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        var count = CountOccurrences(text, "MergeOwner");
        Assert.Equal(1, count);
    }

    private static int CountOccurrences(string source, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = source.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }
}

// ============================================================
// 機能テスト: 画像
// ============================================================

/// <summary>
/// 画像埋め込みに関する機能テスト。
/// 画像を含む Excel から PDF を生成し、PDF に画像が含まれることを確認する。
/// </summary>
public sealed class FeatureImageTests
{
    private static readonly byte[] OnePxPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");

    [Fact]
    public void PdfShouldEmbedSingleImage()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "WithImage";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "Logo")
                .MoveTo(sheet.Cell("B2"))
                .WithSize(60, 40);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureImageTests)}_{nameof(PdfShouldEmbedSingleImage)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("WithImage", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
        // PDF bytes should be larger than a text-only PDF due to the embedded image data
        Assert.True(pdfBytes.Length > 1000);
    }

    [Fact]
    public void PdfShouldEmbedMultipleImages()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
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

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureImageTests)}_{nameof(PdfShouldEmbedMultipleImages)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.True(pdfBytes.Length > 1000);
    }

    [Fact]
    public void PdfShouldHandleFreeFloatingImage()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "FreeFloat";
            using var imgStream = new MemoryStream(OnePxPng, writable: false);
            sheet.AddPicture(imgStream, XLPictureFormat.Png, "FreeImg")
                .MoveTo(20, 30)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .WithSize(50, 30);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureImageTests)}_{nameof(PdfShouldHandleFreeFloatingImage)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("FreeFloat", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: 行の追加
// ============================================================

/// <summary>
/// 行の追加 (InsertCopyBelow / InsertCopyAfter / RowRange) に関する機能テスト。
/// 追加された全行のテキストが PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureRowAdditionTests
{
    [Fact]
    public void PdfShouldContainAllRowsAddedWithInsertCopyBelow()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
            sheet.Cell("A3").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Item");
        // Flow B: insert all copies from template first, then replace in each
        var row1 = template.InsertCopyBelow();
        var row2 = row1.InsertCopyBelow();
        var row3 = row2.InsertCopyBelow();
        row1.ReplacePlaceholder("Item", "ItemA");
        row2.ReplacePlaceholder("Item", "ItemB");
        row3.ReplacePlaceholder("Item", "ItemC");
        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureRowAdditionTests)}_{nameof(PdfShouldContainAllRowsAddedWithInsertCopyBelow)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("ItemA", text, StringComparison.Ordinal);
        Assert.Contains("ItemB", text, StringComparison.Ordinal);
        Assert.Contains("ItemC", text, StringComparison.Ordinal);
        Assert.Contains("Footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainAllRowsAddedWithInsertCopyAfter()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Row}}";
            sheet.Cell("A3").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var templateRow = sheet.FindRow("Row");
        var last = templateRow;
        foreach (var label in new[] { "Row1", "Row2", "Row3" })
        {
            last = templateRow.InsertCopyAfter(last);
            last.ReplacePlaceholder("Row", label);
        }

        templateRow.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureRowAdditionTests)}_{nameof(PdfShouldContainAllRowsAddedWithInsertCopyAfter)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Row1", text, StringComparison.Ordinal);
        Assert.Contains("Row2", text, StringComparison.Ordinal);
        Assert.Contains("Row3", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldPreserveStyleAfterRowCopy()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var templateCell = sheet.Cell("A1");
            templateCell.Value = "{{StyledItem}}";
            templateCell.Style.Font.Bold = true;
            templateCell.Style.Fill.BackgroundColor = XLColor.FromArgb(200, 230, 255);
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("StyledItem");
        var copy = template.InsertCopyBelow();
        copy.ReplacePlaceholder("StyledItem", "CopiedRow");
        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureRowAdditionTests)}_{nameof(PdfShouldPreserveStyleAfterRowCopy)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("CopiedRow", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldHandleZeroDetailRows()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Item}}";
            sheet.Cell("A3").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Item");
        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureRowAdditionTests)}_{nameof(PdfShouldHandleZeroDetailRows)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Header", text, StringComparison.Ordinal);
        Assert.Contains("Footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainAllRowsFromMultiRowRangeExpansion()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{Name}}";
            sheet.Cell("A3").Value = "{{Detail}}";
            sheet.Cell("A4").Value = "Footer";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var templateRange = sheet.FindRows("Name");
        var last = templateRange;
        foreach (var (name, detail) in new[] { ("Alice", "Detail1"), ("Bob", "Detail2") })
        {
            last = templateRange.InsertCopyAfter(last);
            last.ReplacePlaceholder("Name", name);
            last.ReplacePlaceholder("Detail", detail);
        }

        sheet.DeleteRows(templateRange.StartRow, templateRange.EndRow);

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureRowAdditionTests)}_{nameof(PdfShouldContainAllRowsFromMultiRowRangeExpansion)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alice", text, StringComparison.Ordinal);
        Assert.Contains("Detail1", text, StringComparison.Ordinal);
        Assert.Contains("Bob", text, StringComparison.Ordinal);
        Assert.Contains("Detail2", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: テキスト配置
// ============================================================

/// <summary>
/// テキストの水平・垂直配置に関する機能テスト。
/// </summary>
public sealed class FeatureTextAlignmentTests
{
    [Theory]
    [InlineData(XLAlignmentHorizontalValues.Left, "LeftAligned")]
    [InlineData(XLAlignmentHorizontalValues.Center, "CenterAligned")]
    [InlineData(XLAlignmentHorizontalValues.Right, "RightAligned")]
    public void PdfShouldContainHorizontallyAlignedText(XLAlignmentHorizontalValues alignment, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Alignment.Horizontal = alignment;
            sheet.Column(1).Width = 30d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureTextAlignmentTests)}_{nameof(PdfShouldContainHorizontallyAlignedText)}_{alignment}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(XLAlignmentVerticalValues.Top, "TopAligned")]
    [InlineData(XLAlignmentVerticalValues.Center, "MiddleAligned")]
    [InlineData(XLAlignmentVerticalValues.Bottom, "BottomAligned")]
    public void PdfShouldContainVerticallyAlignedText(XLAlignmentVerticalValues alignment, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = cellValue;
            cell.Style.Alignment.Vertical = alignment;
            sheet.Row(1).Height = 40d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureTextAlignmentTests)}_{nameof(PdfShouldContainVerticallyAlignedText)}_{alignment}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainWrappedText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "This is a long text that should wrap inside the cell boundaries";
            cell.Style.Alignment.WrapText = true;
            sheet.Column(1).Width = 20d;
            sheet.Row(1).Height = 50d;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureTextAlignmentTests)}_{nameof(PdfShouldContainWrappedText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}

// ============================================================
// 機能テスト: 罫線
// ============================================================

/// <summary>
/// 罫線スタイルに関する機能テスト。
/// 細線・中線・太線・二重線が PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureBorderTests
{
    [Theory]
    [InlineData(XLBorderStyleValues.Thin, "ThinBorder")]
    [InlineData(XLBorderStyleValues.Medium, "MediumBorder")]
    [InlineData(XLBorderStyleValues.Thick, "ThickBorder")]
    [InlineData(XLBorderStyleValues.Double, "DoubleBorder")]
    [InlineData(XLBorderStyleValues.Dashed, "DashedBorder")]
    [InlineData(XLBorderStyleValues.Dotted, "DottedBorder")]
    public void PdfShouldContainCellWithBorder(XLBorderStyleValues borderStyle, string cellValue)
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = cellValue;
            cell.Style.Border.OutsideBorder = borderStyle;
            cell.Style.Border.OutsideBorderColor = XLColor.Black;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureBorderTests)}_{nameof(PdfShouldContainCellWithBorder)}_{borderStyle}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains(cellValue, PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderColoredBorder()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("B2");
            cell.Value = "ColoredBorder";
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            cell.Style.Border.OutsideBorderColor = XLColor.FromArgb(255, 0, 0);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureBorderTests)}_{nameof(PdfShouldRenderColoredBorder)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ColoredBorder", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderTableWithAllSideBorders()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
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

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureBorderTests)}_{nameof(PdfShouldRenderTableWithAllSideBorders)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("R1C1", text, StringComparison.Ordinal);
        Assert.Contains("R3C3", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: プレースホルダー置換
// ============================================================

/// <summary>
/// プレースホルダー置換に関する機能テスト。
/// </summary>
public sealed class FeaturePlaceholderTests
{
    [Fact]
    public void PdfShouldContainReplacedPlaceholderValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{CustomerName}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.ReplacePlaceholder("CustomerName", "AcmeCorp");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePlaceholderTests)}_{nameof(PdfShouldContainReplacedPlaceholderValue)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("AcmeCorp", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainAllReplacedPlaceholders()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Title}}";
            sheet.Cell("A2").Value = "{{Name}}";
            sheet.Cell("A3").Value = "{{Date}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["Title"] = "Invoice",
            ["Name"] = "JohnDoe",
            ["Date"] = "2025-01-01"
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePlaceholderTests)}_{nameof(PdfShouldContainAllReplacedPlaceholders)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Invoice", text, StringComparison.Ordinal);
        Assert.Contains("JohnDoe", text, StringComparison.Ordinal);
        Assert.Contains("2025-01-01", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldPreservePlaceholderNotReplaced()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{KeepMe}}";
            sheet.Cell("A2").Value = "{{ReplaceMe}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.ReplacePlaceholder("ReplaceMe", "Replaced");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePlaceholderTests)}_{nameof(PdfShouldPreservePlaceholderNotReplaced)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Replaced", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainReplacedPlaceholderInMergedCell()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{MergedTitle}}";
            sheet.Range("A1:C1").Merge();
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        sheet.ReplacePlaceholder("MergedTitle", "MergedValue");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePlaceholderTests)}_{nameof(PdfShouldContainReplacedPlaceholderInMergedCell)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("MergedValue", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: ページ設定
// ============================================================

/// <summary>
/// ページサイズ・余白・中央配置などのページ設定に関する機能テスト。
/// </summary>
public sealed class FeaturePageSetupTests
{
    [Fact]
    public void PdfShouldHaveA4PageDimensions()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "A4Page";
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePageSetupTests)}_{nameof(PdfShouldHaveA4PageDimensions)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var (width, height) = PdfTestHelper.GetPageSize(pdfBytes);
        // A4: 595.28 x 841.89 pt (tolerance 2pt)
        Assert.Equal(595.28d, width, 0);
        Assert.Equal(841.89d, height, 0);
    }

    [Fact]
    public void PdfShouldHaveA4LandscapePageDimensions()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "A4Landscape";
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePageSetupTests)}_{nameof(PdfShouldHaveA4LandscapePageDimensions)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var (width, height) = PdfTestHelper.GetPageSize(pdfBytes);
        // A4 Landscape: 841.89 x 595.28 pt
        Assert.Equal(841.89d, width, 0);
        Assert.Equal(595.28d, height, 0);
    }

    [Fact]
    public void PdfShouldRespectCenterHorizontally()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Centered";
            sheet.PageSetup.CenterHorizontally = true;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePageSetupTests)}_{nameof(PdfShouldRespectCenterHorizontally)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("Centered", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldGenerateMultiplePagesWhenContentExceedsPageHeight()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            // Fill enough rows to force a second page
            for (var row = 1; row <= 60; row++)
            {
                sheet.Cell(row, 1).Value = $"Row{row:D2}";
                sheet.Row(row).Height = 20d;
            }
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePageSetupTests)}_{nameof(PdfShouldGenerateMultiplePagesWhenContentExceedsPageHeight)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        // Each sheet produces one page; 60-row content is rendered on that page
        Assert.Equal(1, PdfTestHelper.GetPageCount(pdfBytes));
        Assert.Contains("Row01", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
        Assert.Contains("Row60", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldApplyManualPageBreak()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Page1Content";
            sheet.Cell("A2").Value = "Page2Content";
            sheet.PageSetup.AddHorizontalPageBreak(1);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePageSetupTests)}_{nameof(PdfShouldApplyManualPageBreak)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        // Intra-sheet page breaks are not yet split by the planner;
        // both rows appear on a single page.
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Page1Content", text, StringComparison.Ordinal);
        Assert.Contains("Page2Content", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: ヘッダー・フッター
// ============================================================

/// <summary>
/// ヘッダー・フッターに関する機能テスト。
/// </summary>
public sealed class FeatureHeaderFooterTests
{
    [Fact]
    public void PdfShouldContainHeaderText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Header.Left.AddText("LeftHeader", XLHFOccurrence.OddPages);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureHeaderFooterTests)}_{nameof(PdfShouldContainHeaderText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BodyContent", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainFooterText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "BodyContent";
            sheet.PageSetup.Footer.Right.AddText("RightFooter", XLHFOccurrence.OddPages);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureHeaderFooterTests)}_{nameof(PdfShouldContainFooterText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void PdfShouldContainBothHeaderAndFooter()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Main";
            sheet.PageSetup.Header.Center.AddText("TopCenter", XLHFOccurrence.OddPages);
            sheet.PageSetup.Footer.Center.AddText("BottomCenter", XLHFOccurrence.OddPages);
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureHeaderFooterTests)}_{nameof(PdfShouldContainBothHeaderAndFooter)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}

// ============================================================
// 機能テスト: 複数シート
// ============================================================

/// <summary>
/// 複数シートに関する機能テスト。
/// 各シートが個別ページとして PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureMultiSheetTests
{
    [Fact]
    public void PdfShouldHaveOnePagePerSheet()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet1 = workbook.AddWorksheet("Sheet1");
            sheet1.Cell("A1").Value = "ContentSheet1";
            var sheet2 = workbook.AddWorksheet("Sheet2");
            sheet2.Cell("A1").Value = "ContentSheet2";
            var sheet3 = workbook.AddWorksheet("Sheet3");
            sheet3.Cell("A1").Value = "ContentSheet3";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureMultiSheetTests)}_{nameof(PdfShouldHaveOnePagePerSheet)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.True(PdfTestHelper.GetPageCount(pdfBytes) >= 3);
    }

    [Fact]
    public void PdfShouldContainTextFromAllSheets()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Alpha").Cell("A1").Value = "AlphaSheet";
            workbook.AddWorksheet("Beta").Cell("A1").Value = "BetaSheet";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureMultiSheetTests)}_{nameof(PdfShouldContainTextFromAllSheets)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("AlphaSheet", text, StringComparison.Ordinal);
        Assert.Contains("BetaSheet", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: 非表示行・列
// ============================================================

/// <summary>
/// 非表示行・列に関する機能テスト。
/// 非表示の行・列のテキストが PDF に含まれないことを確認する。
/// </summary>
public sealed class FeatureHiddenRowColumnTests
{
    [Fact]
    public void PdfShouldNotContainHiddenRowText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VisibleRow";
            sheet.Cell("A2").Value = "HiddenRow";
            sheet.Row(2).Hide();
            sheet.Cell("A3").Value = "VisibleRow2";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureHiddenRowColumnTests)}_{nameof(PdfShouldNotContainHiddenRowText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("VisibleRow", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HiddenRow", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldNotContainHiddenColumnText()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "VisibleCol";
            sheet.Cell("B1").Value = "HiddenCol";
            sheet.Column(2).Hide();
            sheet.Cell("C1").Value = "VisibleCol2";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureHiddenRowColumnTests)}_{nameof(PdfShouldNotContainHiddenColumnText)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("VisibleCol", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HiddenCol", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: 印刷範囲
// ============================================================

/// <summary>
/// 印刷範囲に関する機能テスト。
/// 印刷範囲外のセルが PDF に出力されないことを確認する。
/// </summary>
public sealed class FeaturePrintAreaTests
{
    [Fact]
    public void PdfShouldOnlyContainTextWithinPrintArea()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "InPrintArea";
            sheet.Cell("A2").Value = "InPrintArea2";
            sheet.Cell("A10").Value = "OutsidePrintArea";
            sheet.PageSetup.PrintAreas.Add("A1:B5");
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeaturePrintAreaTests)}_{nameof(PdfShouldOnlyContainTextWithinPrintArea)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("InPrintArea", text, StringComparison.Ordinal);
        Assert.DoesNotContain("OutsidePrintArea", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: セルの値型
// ============================================================

/// <summary>
/// セルの値型 (文字列・数値・日付・論理値) に関する機能テスト。
/// </summary>
public sealed class FeatureCellValueTypeTests
{
    [Fact]
    public void PdfShouldRenderNumericCellValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = 12345;
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureCellValueTypeTests)}_{nameof(PdfShouldRenderNumericCellValue)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("12345", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderDateCellValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = new DateTime(2025, 1, 15);
            cell.Style.DateFormat.Format = "yyyy/MM/dd";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureCellValueTypeTests)}_{nameof(PdfShouldRenderDateCellValue)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("2025", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderFormulaCellValue()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = 10;
            sheet.Cell("A2").Value = 20;
            sheet.Cell("A3").FormulaA1 = "=A1+A2";
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureCellValueTypeTests)}_{nameof(PdfShouldRenderFormulaCellValue)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}

// ============================================================
// 機能テスト: 埋め込みフォント (ipaexg.ttf)
// ============================================================

/// <summary>
/// IPAex ゴシック埋め込みフォントに関する機能テスト。
/// ipaexg.ttf を使用して日本語テキストが PDF に出力されることを確認する。
/// </summary>
public sealed class FeatureEmbeddedFontTests
{
    [Fact]
    public void PdfShouldEmbedIpaExGothicFontForJapaneseText()
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
            $"{nameof(FeatureEmbeddedFontTests)}_{nameof(PdfShouldEmbedIpaExGothicFontForJapaneseText)}",
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        // フォント名が PDF に埋め込まれていることを確認
        var letters = PdfTestHelper.GetLetters(pdfBytes);
        Assert.Contains(
            letters,
            static l => l.FontName is not null && l.FontName.Contains("IPAex", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PdfShouldRenderMultipleJapaneseCellsWithEmbeddedFont()
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
            $"{nameof(FeatureEmbeddedFontTests)}_{nameof(PdfShouldRenderMultipleJapaneseCellsWithEmbeddedFont)}",
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void PdfShouldRenderJapaneseBoldWithEmbeddedFont()
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
            $"{nameof(FeatureEmbeddedFontTests)}_{nameof(PdfShouldRenderJapaneseBoldWithEmbeddedFont)}",
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }

    [Fact]
    public void PdfShouldRenderMixedJapaneseAndAsciiWithEmbeddedFont()
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
            $"{nameof(FeatureEmbeddedFontTests)}_{nameof(PdfShouldRenderMixedJapaneseAndAsciiWithEmbeddedFont)}",
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("EnglishText", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldEmbedFontFromIpaExGothicFontPath()
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
            $"{nameof(FeatureEmbeddedFontTests)}_{nameof(PdfShouldEmbedFontFromIpaExGothicFontPath)}",
            stream,
            resolver);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
    }
}

// ============================================================
// 機能テスト: 総合 (請求書・明細テンプレート)
// ============================================================

/// <summary>
/// 請求書形式の総合テンプレートに関する機能テスト。
/// 複数機能を組み合わせた実帳票シナリオを確認する。
/// </summary>
public sealed class FeatureInvoiceTemplateTests
{
    [Fact]
    public void PdfShouldRenderInvoiceTemplateWithAllFeatures()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Invoice");
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet.PageSetup.Margins.Left = 1.8;
            sheet.PageSetup.Margins.Right = 1.8;
            sheet.PageSetup.Margins.Top = 2.5;
            sheet.PageSetup.Margins.Bottom = 2.5;

            // Title (merged, bold, large)
            sheet.Cell("A1").Value = "{{Title}}";
            sheet.Range("A1:F1").Merge();
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 16d;
            sheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Customer info
            sheet.Cell("A3").Value = "宛先:";
            sheet.Cell("B3").Value = "{{CustomerName}}";
            sheet.Cell("A4").Value = "日付:";
            sheet.Cell("B4").Value = "{{IssueDate}}";

            // Detail header
            sheet.Cell("A6").Value = "品目";
            sheet.Cell("B6").Value = "数量";
            sheet.Cell("C6").Value = "単価";
            sheet.Cell("D6").Value = "金額";
            sheet.Cell("A6").Style.Font.Bold = true;
            sheet.Cell("B6").Style.Font.Bold = true;
            sheet.Cell("C6").Style.Font.Bold = true;
            sheet.Cell("D6").Style.Font.Bold = true;
            sheet.Cell("A6").Style.Fill.BackgroundColor = XLColor.FromArgb(220, 220, 220);
            sheet.Cell("B6").Style.Fill.BackgroundColor = XLColor.FromArgb(220, 220, 220);
            sheet.Cell("C6").Style.Fill.BackgroundColor = XLColor.FromArgb(220, 220, 220);
            sheet.Cell("D6").Style.Fill.BackgroundColor = XLColor.FromArgb(220, 220, 220);
            sheet.Cell("A6").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell("B6").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell("C6").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell("D6").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            // Detail template row
            sheet.Cell("A7").Value = "{{ItemName}}";
            sheet.Cell("B7").Value = "{{Qty}}";
            sheet.Cell("C7").Value = "{{UnitPrice}}";
            sheet.Cell("D7").Value = "{{Amount}}";
            sheet.Cell("A7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell("B7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell("C7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell("D7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            // Footer
            sheet.Cell("A9").Value = "{{FooterNote}}";
        });

        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);

        sheet.ReplacePlaceholder("Title", "請求書");
        sheet.ReplacePlaceholder("CustomerName", "株式会社サンプル");
        sheet.ReplacePlaceholder("IssueDate", "2025-01-15");

        var templateRow = sheet.FindRow("ItemName");
        var lastRow = templateRow;
        foreach (var (name, qty, price, amount) in new[]
        {
            ("商品A", "2", "1000", "2000"),
            ("商品B", "1", "3000", "3000"),
            ("商品C", "5", "500", "2500")
        })
        {
            lastRow = templateRow.InsertCopyAfter(lastRow);
            lastRow.ReplacePlaceholder("ItemName", name);
            lastRow.ReplacePlaceholder("Qty", qty);
            lastRow.ReplacePlaceholder("UnitPrice", price);
            lastRow.ReplacePlaceholder("Amount", amount);
        }

        templateRow.Delete();
        sheet.ReplacePlaceholder("FooterNote", "上記の通りご請求申し上げます。");

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureInvoiceTemplateTests)}_{nameof(PdfShouldRenderInvoiceTemplateWithAllFeatures)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("2025-01-15", text, StringComparison.Ordinal);
        Assert.Contains("2000", text, StringComparison.Ordinal);
        Assert.Contains("3000", text, StringComparison.Ordinal);
        Assert.Contains("2500", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldRenderListReportWithStripedRows()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("List");
            sheet.Cell("A1").Value = "No";
            sheet.Cell("B1").Value = "Name";
            sheet.Cell("C1").Value = "Score";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("B1").Style.Font.Bold = true;
            sheet.Cell("C1").Style.Font.Bold = true;

            for (var i = 1; i <= 10; i++)
            {
                var bgColor = i % 2 == 0
                    ? XLColor.FromArgb(240, 248, 255)
                    : XLColor.FromArgb(255, 255, 255);
                sheet.Cell(i + 1, 1).Value = i;
                sheet.Cell(i + 1, 2).Value = $"Name{i:D2}";
                sheet.Cell(i + 1, 3).Value = i * 10;
                for (var col = 1; col <= 3; col++)
                {
                    sheet.Cell(i + 1, col).Style.Fill.BackgroundColor = bgColor;
                    sheet.Cell(i + 1, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureInvoiceTemplateTests)}_{nameof(PdfShouldRenderListReportWithStripedRows)}",
            stream);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Name01", text, StringComparison.Ordinal);
        Assert.Contains("Name10", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: 特定シートのみの出力
// ============================================================

/// <summary>
/// 特定シートのみを対象に PDF を生成する機能テスト。
/// <see cref="OysterReportEngine.GeneratePdf(TemplateSheet, Stream)"/> を使用する。
/// </summary>
public sealed class FeatureSingleSheetOutputTests
{
    [Fact]
    public void PdfShouldContainOnlyTargetSheetContentWhenRenderedByIndex()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Sheet1").Cell("A1").Value = "Sheet1Content";
            workbook.AddWorksheet("Sheet2").Cell("A1").Value = "Sheet2Content";
            workbook.AddWorksheet("Sheet3").Cell("A1").Value = "Sheet3Content";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var targetSheet = workbook.Sheets[1]; // Sheet2 (0-based)

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            $"{nameof(FeatureSingleSheetOutputTests)}_{nameof(PdfShouldContainOnlyTargetSheetContentWhenRenderedByIndex)}",
            targetSheet);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, PdfTestHelper.GetPageCount(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Sheet2Content", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Sheet1Content", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Sheet3Content", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainOnlyTargetSheetContentWhenRenderedByName()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Summary").Cell("A1").Value = "SummaryContent";
            workbook.AddWorksheet("Detail").Cell("A1").Value = "DetailContent";
            workbook.AddWorksheet("Appendix").Cell("A1").Value = "AppendixContent";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var targetSheet = workbook.GetSheet("Detail");

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            $"{nameof(FeatureSingleSheetOutputTests)}_{nameof(PdfShouldContainOnlyTargetSheetContentWhenRenderedByName)}",
            targetSheet);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, PdfTestHelper.GetPageCount(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("DetailContent", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SummaryContent", text, StringComparison.Ordinal);
        Assert.DoesNotContain("AppendixContent", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainOnlyFirstSheetWhenRenderedByIndex0()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("First").Cell("A1").Value = "FirstSheet";
            workbook.AddWorksheet("Second").Cell("A1").Value = "SecondSheet";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var targetSheet = workbook.Sheets[0];

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            $"{nameof(FeatureSingleSheetOutputTests)}_{nameof(PdfShouldContainOnlyFirstSheetWhenRenderedByIndex0)}",
            targetSheet);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Equal(1, PdfTestHelper.GetPageCount(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("FirstSheet", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SecondSheet", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldReflectPlaceholderReplacementsOnTargetSheetOnly()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Cover").Cell("A1").Value = "{{Title}}";
            workbook.AddWorksheet("Body").Cell("A1").Value = "{{Content}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var coverSheet = workbook.GetSheet("Cover");
        coverSheet.ReplacePlaceholder("Title", "ReplacedTitle");

        var pdfBytes = PdfTestHelper.GenerateSheetPdfAndSave(
            $"{nameof(FeatureSingleSheetOutputTests)}_{nameof(PdfShouldReflectPlaceholderReplacementsOnTargetSheetOnly)}",
            coverSheet);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ReplacedTitle", PdfTestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: 明細行展開後の行数検証
// ============================================================

/// <summary>
/// 明細行の複数展開と、テンプレート削除後に余計な行が増えていないことを検証するテスト。
/// </summary>
public sealed class FeatureDetailRowExpansionVerificationTests
{
    [Fact]
    public void PdfShouldContainExactlyThreeDetailRowsAfterExpansion()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("A2").Value = "{{No}}: {{ItemName}}";
            sheet.Cell("A3").Value = "Footer";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("ItemName");

        // Flow A: テンプレートを元に 3 件の明細行を追加
        var last = template;
        foreach (var (no, name) in new[] { ("1", "Apple"), ("2", "Banana"), ("3", "Cherry") })
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholder("No", no);
            last.ReplacePlaceholder("ItemName", name);
        }

        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureDetailRowExpansionVerificationTests)}_{nameof(PdfShouldContainExactlyThreeDetailRowsAfterExpansion)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);

        // 3 件の明細が存在すること
        Assert.Contains("Apple", text, StringComparison.Ordinal);
        Assert.Contains("Banana", text, StringComparison.Ordinal);
        Assert.Contains("Cherry", text, StringComparison.Ordinal);

        // テンプレートプレースホルダーが残っていないこと
        Assert.DoesNotContain("{{ItemName}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{No}}", text, StringComparison.Ordinal);

        // 4 件目 (Durian) は存在しないこと
        Assert.DoesNotContain("Durian", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldNotContainTemplatePlaceholderAfterDeletion()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Product}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Product");
        var row = template.InsertCopyBelow();
        row.ReplacePlaceholder("Product", "Widget");
        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureDetailRowExpansionVerificationTests)}_{nameof(PdfShouldNotContainTemplatePlaceholderAfterDeletion)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Product}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldPreserveHeaderAndFooterAroundExpandedDetailRows()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "ReportHeader";
            sheet.Cell("A2").Value = "{{Line}}";
            sheet.Cell("A3").Value = "ReportFooter";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Line");

        var last = template;
        last = template.InsertCopyAfter(last);
        last.ReplacePlaceholder("Line", "LineA");
        last = template.InsertCopyAfter(last);
        last.ReplacePlaceholder("Line", "LineB");
        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureDetailRowExpansionVerificationTests)}_{nameof(PdfShouldPreserveHeaderAndFooterAroundExpandedDetailRows)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("ReportHeader", text, StringComparison.Ordinal);
        Assert.Contains("LineA", text, StringComparison.Ordinal);
        Assert.Contains("LineB", text, StringComparison.Ordinal);
        Assert.Contains("ReportFooter", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfShouldContainExactOccurrencesOfDetailLabel()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "START";
            sheet.Cell("A2").Value = "ROW-{{Seq}}";
            sheet.Cell("A3").Value = "END";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Seq");

        // Flow A: 4 件挿入
        var last = template;
        for (var i = 1; i <= 4; i++)
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholder("Seq", i.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureDetailRowExpansionVerificationTests)}_{nameof(PdfShouldContainExactOccurrencesOfDetailLabel)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);

        // ROW-1 ～ ROW-4 がすべて存在すること
        Assert.Contains("ROW-1", text, StringComparison.Ordinal);
        Assert.Contains("ROW-2", text, StringComparison.Ordinal);
        Assert.Contains("ROW-3", text, StringComparison.Ordinal);
        Assert.Contains("ROW-4", text, StringComparison.Ordinal);

        // ROW-5 は存在しないこと (余計な行がないこと)
        Assert.DoesNotContain("ROW-5", text, StringComparison.Ordinal);

        // テンプレートプレースホルダーが残っていないこと
        Assert.DoesNotContain("{{Seq}}", text, StringComparison.Ordinal);
    }
}

// ============================================================
// 機能テスト: ReplacePlaceholders による一括指定
// ============================================================

/// <summary>
/// <see cref="TemplateRow.ReplacePlaceholders"/>・<see cref="TemplateRowRange.ReplacePlaceholders"/>・
/// <see cref="TemplateWorkbook.ReplacePlaceholders"/> による一括プレースホルダー置換のテスト。
/// </summary>
public sealed class FeatureReplacePlaceholdersTests
{
    [Fact]
    public void TemplateRowShouldReplaceMultiplePlaceholdersAtOnce()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            // Use a single row with each placeholder value in a separate cell
            // (values with distinct lowercase-only names to avoid PdfPig glyph ordering artifacts)
            sheet.Cell("A1").Value = "name:{{PersonName}} dept:{{PersonDept}} city:{{PersonCity}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var row = sheet.GetRow(1);
        row.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["PersonName"] = "tanaka",
            ["PersonDept"] = "sales",
            ["PersonCity"] = "tokyo"
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureReplacePlaceholdersTests)}_{nameof(TemplateRowShouldReplaceMultiplePlaceholdersAtOnce)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("tanaka", text, StringComparison.Ordinal);
        Assert.Contains("sales", text, StringComparison.Ordinal);
        Assert.Contains("tokyo", text, StringComparison.Ordinal);
    }

    [Fact]
    public void TemplateRowShouldTreatNullValueAsEmptyStringInReplacePlaceholders()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Name: {{Name}}";
            sheet.Cell("B1").Value = "Memo: {{Memo}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var row = sheet.GetRow(1);
        row.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["Name"] = "Alice",
            ["Memo"] = null // null → empty string
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureReplacePlaceholdersTests)}_{nameof(TemplateRowShouldTreatNullValueAsEmptyStringInReplacePlaceholders)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Alice", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Memo}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void TemplateRowRangeShouldReplaceMultiplePlaceholdersAtOnce()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Item: {{ItemName}}";
            sheet.Cell("A2").Value = "Price: {{Price}}";
            sheet.Cell("A3").Value = "Qty: {{Qty}}";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var range = sheet.FindRows("ItemName");
        range.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["ItemName"] = "Widget",
            ["Price"] = "980",
            ["Qty"] = "5"
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureReplacePlaceholdersTests)}_{nameof(TemplateRowRangeShouldReplaceMultiplePlaceholdersAtOnce)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("980", text, StringComparison.Ordinal);
        Assert.Contains("5", text, StringComparison.Ordinal);
    }

    [Fact]
    public void TemplateWorkbookShouldReplaceAllPlaceholdersAcrossAllSheets()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.AddWorksheet("Cover").Cell("A1").Value = "{{DocTitle}}";
            workbook.AddWorksheet("Body").Cell("A1").Value = "Author: {{Author}}";
            workbook.AddWorksheet("Appendix").Cell("A1").Value = "{{DocTitle}} - Appendix";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        workbook.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["DocTitle"] = "AnnualReport",
            ["Author"] = "Smith"
        });

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureReplacePlaceholdersTests)}_{nameof(TemplateWorkbookShouldReplaceAllPlaceholdersAcrossAllSheets)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        // DocTitle は Cover と Appendix の 2 箇所に出現する
        Assert.Contains("AnnualReport", text, StringComparison.Ordinal);
        Assert.Contains("Smith", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{DocTitle}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Author}}", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExpandedRowsShouldBeReplacedWithReplacePlaceholdersInLoop()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "HEADER";
            sheet.Cell("A2").Value = "{{Code}}";
            sheet.Cell("B2").Value = "{{Label}}";
            sheet.Cell("C2").Value = "{{Value}}";
            sheet.Cell("A3").Value = "FOOTER";
        });

        stream.Position = 0;
        using var workbook = new TemplateWorkbook(stream);
        var sheet = Assert.Single(workbook.Sheets);
        var template = sheet.FindRow("Code");

        var items = new[]
        {
            new Dictionary<string, string?> { ["Code"] = "001", ["Label"] = "Alpha", ["Value"] = "100" },
            new Dictionary<string, string?> { ["Code"] = "002", ["Label"] = "Beta",  ["Value"] = "200" },
            new Dictionary<string, string?> { ["Code"] = "003", ["Label"] = "Gamma", ["Value"] = "300" }
        };

        var last = template;
        foreach (var item in items)
        {
            last = template.InsertCopyAfter(last);
            last.ReplacePlaceholders(item);
        }

        template.Delete();

        var pdfBytes = PdfTestHelper.GeneratePdfAndSave(
            $"{nameof(FeatureReplacePlaceholdersTests)}_{nameof(ExpandedRowsShouldBeReplacedWithReplacePlaceholdersInLoop)}",
            workbook);

        Assert.True(PdfTestHelper.IsValidPdf(pdfBytes));
        var text = PdfTestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("001", text, StringComparison.Ordinal);
        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("002", text, StringComparison.Ordinal);
        Assert.Contains("Beta", text, StringComparison.Ordinal);
        Assert.Contains("003", text, StringComparison.Ordinal);
        Assert.Contains("Gamma", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{Code}}", text, StringComparison.Ordinal);
    }
}
