namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Tests.Helpers;

using Xunit;

public sealed partial class FeatureTests
{
    [Fact]
    public void FontStyleShouldRenderBoldText()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldText";
            cell.Style.Font.Bold = true;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderBoldText), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldText", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderItalicText()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "ItalicText";
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderItalicText), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("ItalicText", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderBoldItalicText()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "BoldItalic";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Italic = true;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderBoldItalicText), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("BoldItalic", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderMixedStylesOnSamePage()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
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

        var pdfBytes = TestHelper.GeneratePdfAndSave(
            nameof(FontStyleShouldRenderMixedStylesOnSamePage),
            stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        var text = TestHelper.ExtractAllText(pdfBytes);
        Assert.Contains("Normal", text, StringComparison.Ordinal);
        Assert.Contains("Bold", text, StringComparison.Ordinal);
        Assert.Contains("Italic", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderUnderlinedText()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "UnderlinedText";
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderUnderlinedText), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("UnderlinedText", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderStrikethroughText()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "StrikethroughText";
            cell.Style.Font.Strikethrough = true;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderStrikethroughText), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("StrikethroughText", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }

    [Fact]
    public void FontStyleShouldRenderAllDecorationsOnSameCell()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "AllDecorations";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Italic = true;
            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
            cell.Style.Font.Strikethrough = true;
        });

        var pdfBytes = TestHelper.GeneratePdfAndSave(nameof(FontStyleShouldRenderAllDecorationsOnSameCell), stream);

        Assert.True(TestHelper.IsValidPdf(pdfBytes));
        Assert.Contains("AllDecorations", TestHelper.ExtractAllText(pdfBytes), StringComparison.Ordinal);
    }
}
