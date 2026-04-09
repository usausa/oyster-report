// <copyright file="PdfGeneratorTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Internal;

using Xunit;

public sealed class PdfGeneratorTests
{
    [Fact]
    public void GenerateShouldCreatePdfDocument()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
            sheet.Cell("A2").Value = "World";
        });

        var workbook = ExcelReader.Read(stream);
        var renderPlan = PdfRenderPlanner.BuildPlan(workbook);
        using var output = new MemoryStream();

        var context = new ReportRenderContext
        {
            Workbook = workbook,
            SheetPlans = renderPlan
        };

        PdfGenerator.WritePdf(context, output);

        output.Position = 0;
        using var reader = new StreamReader(output, leaveOpen: true);
        var header = reader.ReadLine();
        Assert.NotNull(header);
        Assert.StartsWith("%PDF", header, StringComparison.Ordinal);
    }

    [Fact]
    public void GenerateShouldPassBoldAndItalicRequestsToFontResolver()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");

            var boldCell = sheet.Cell("A1");
            boldCell.Value = "Bold";
            boldCell.Style.Font.FontName = "CustomEmbeddedFont";
            boldCell.Style.Font.Bold = true;

            var italicCell = sheet.Cell("A2");
            italicCell.Value = "Italic";
            italicCell.Style.Font.FontName = "CustomEmbeddedFont";
            italicCell.Style.Font.Italic = true;
        });

        using var template = new TemplateWorkbook(stream);
        var resolver = new TrackingEmbeddedFontResolver(GetEmbeddedFontBytes());
        var engine = new OysterReportEngine { FontResolver = resolver };
        using var output = new MemoryStream();

        engine.GeneratePdf(template, output);

        Assert.Contains(resolver.Requests, request => request.FamilyName == "CustomEmbeddedFont" && request.Bold);
        Assert.Contains(resolver.Requests, request => request.FamilyName == "CustomEmbeddedFont" && request.Italic);
    }

    [Fact]
    public void CreateRenderContextShouldKeepLayoutIndependentFromFontResolver()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            workbook.Style.Font.FontName = "Arial";
            workbook.Style.Font.FontSize = 10d;

            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
            sheet.Column(1).Width = 12d;
            sheet.PageSetup.CenterHorizontally = true;
        });

        using var template1 = new TemplateWorkbook(stream);
        var baselineEngine = new OysterReportEngine();
        var baselineContext = baselineEngine.CreateRenderContext(template1);

        stream.Position = 0;
        using var template2 = new TemplateWorkbook(stream);
        var resolvedEngine = new OysterReportEngine
        {
            FontResolver = new EmbeddedFontResolver(GetEmbeddedFontBytes())
        };
        var resolvedContext = resolvedEngine.CreateRenderContext(template2);

        var baselinePage = Assert.Single(baselineContext.SheetPlans[0].Pages);
        var resolvedPage = Assert.Single(resolvedContext.SheetPlans[0].Pages);

        Assert.Equal(baselinePage.PrintableBounds.X, resolvedPage.PrintableBounds.X, 3);
        Assert.Equal(baselinePage.PrintableBounds.Width, resolvedPage.PrintableBounds.Width, 3);
        Assert.Equal(baselinePage.Cells[0].OuterBounds.X, resolvedPage.Cells[0].OuterBounds.X, 3);
        Assert.Equal(baselinePage.Cells[0].OuterBounds.Width, resolvedPage.Cells[0].OuterBounds.Width, 3);
    }

    [Fact]
    public void BuildRenderPlanShouldUseConfiguredPageSizeResolver()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
        });

        var workbook = ExcelReader.Read(stream);
        var options = new ReportRenderingOptions
        {
            PageSizeResolver = static paperSize => paperSize == XLPaperSize.A4Paper ? (700d, 900d) : (595.28d, 841.89d)
        };

        var renderPlan = PdfRenderPlanner.BuildPlan(workbook, options);
        var page = Assert.Single(renderPlan[0].Pages);

        Assert.Equal(700d, page.PageBounds.Width, 3);
        Assert.Equal(900d, page.PageBounds.Height, 3);
    }

    [Fact]
    public void BuildRenderPlanShouldPreserveCellHeightForTextLayout()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Row(1).Height = 9.95d;
            var cell = sheet.Cell("A1");
            cell.Value = "担当：";
            cell.Style.Font.FontName = "ＭＳ Ｐゴシック";
            cell.Style.Font.FontSize = 10d;
            cell.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
        });

        var workbook = ExcelReader.Read(stream);
        var renderPlan = PdfRenderPlanner.BuildPlan(workbook);
        var cell = renderPlan[0].Pages[0].Cells.Single(info => info.CellAddress == "A1");

        Assert.Equal(cell.OuterBounds.Height, cell.ContentBounds.Height, 3);
    }

    [Fact]
    public void DebugDumperShouldWriteWorkbookAndPdfPreparation()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
        });

        using var template = new TemplateWorkbook(stream);
        var engine = new OysterReportEngine();
        var context = engine.CreateRenderContext(template);
        var dumper = new ReportDebugDumper();

        using var workbookDump = new MemoryStream();
        dumper.DumpWorkbook(context, workbookDump);
        var workbookJson = System.Text.Encoding.UTF8.GetString(workbookDump.ToArray());
        Assert.Contains("\"Sheets\"", workbookJson, StringComparison.Ordinal);

        using var pdfPreparationDump = new MemoryStream();
        dumper.DumpPdfPreparation(context, pdfPreparationDump);
        var preparationJson = System.Text.Encoding.UTF8.GetString(pdfPreparationDump.ToArray());
        Assert.Contains("\"RenderPlan\"", preparationJson, StringComparison.Ordinal);
    }

    [Fact]
    public void GenerateShouldSupportEmbeddedFontDataFromResolver()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            var cell = sheet.Cell("A1");
            cell.Value = "埋め込みフォント";
            cell.Style.Font.FontName = "CustomEmbeddedFont";
            cell.Style.Font.FontSize = 10d;
        });

        using var template = new TemplateWorkbook(stream);
        var engine = new OysterReportEngine
        {
            FontResolver = new EmbeddedFontResolver(GetEmbeddedFontBytes())
        };
        using var output = new MemoryStream();

        engine.GeneratePdf(template, output);

        output.Position = 0;
        using var reader = new StreamReader(output, leaveOpen: true);
        var header = reader.ReadLine();
        Assert.NotNull(header);
        Assert.StartsWith("%PDF", header, StringComparison.Ordinal);
    }

    private static byte[] GetEmbeddedFontBytes()
    {
        var fontPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Example", "ipaexg.ttf"));
        return File.ReadAllBytes(fontPath);
    }

    private sealed class EmbeddedFontResolver : IReportFontResolver
    {
        private readonly ReadOnlyMemory<byte> fontData;

        public EmbeddedFontResolver(ReadOnlyMemory<byte> fontData)
        {
            this.fontData = fontData;
        }

        public FontInfo? ResolveTypeface(string familyName, bool bold, bool italic) =>
            string.Equals(familyName, "CustomEmbeddedFont", StringComparison.Ordinal)
                ? new FontInfo { FaceName = "IPAexGothic" }
                : null;

        public ReadOnlyMemory<byte>? GetFontData(string faceName) =>
            string.Equals(faceName, "IPAexGothic", StringComparison.Ordinal)
                ? fontData
                : null;
    }

    private sealed class TrackingEmbeddedFontResolver : IReportFontResolver
    {
        private readonly ReadOnlyMemory<byte> fontData;

        public TrackingEmbeddedFontResolver(ReadOnlyMemory<byte> fontData)
        {
            this.fontData = fontData;
        }

        public List<(string FamilyName, bool Bold, bool Italic)> Requests { get; } = [];

        public FontInfo? ResolveTypeface(string familyName, bool bold, bool italic)
        {
            Requests.Add((familyName, bold, italic));
            return string.Equals(familyName, "CustomEmbeddedFont", StringComparison.Ordinal)
                ? new FontInfo { FaceName = "IPAexGothic" }
                : null;
        }

        public ReadOnlyMemory<byte>? GetFontData(string faceName) =>
            string.Equals(faceName, "IPAexGothic", StringComparison.Ordinal)
                ? fontData
                : null;
    }
}
