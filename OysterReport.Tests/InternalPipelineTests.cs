namespace OysterReport.Tests;

using ClosedXML.Excel;

using OysterReport.Internal;
using OysterReport.Tests.Helpers;

using Xunit;

/// <summary>
/// Excel 読み込み → レンダープラン構築 → PDF 生成の内部パイプラインを直接テストする。
/// フォントリゾルバーとのインターフェース契約・レイアウト計算など、
/// 機能テストが網羅しない内部動作を検証する。
/// </summary>
public sealed class InternalPipelineTests
{
    /// <summary>
    /// <see cref="PdfGenerator.WritePdf"/> が有効な PDF バイナリを出力することを確認する。
    /// </summary>
    [Fact]
    public void WritePdfShouldProduceValidPdfBinary()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
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

    /// <summary>
    /// フォントリゾルバーに Bold/Italic フラグが正しく渡されることを確認する。
    /// </summary>
    [Fact]
    public void GeneratePdfShouldPassBoldAndItalicFlagsToFontResolver()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
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

    /// <summary>
    /// フォントリゾルバーの有無にかかわらず、ページレイアウト (座標・幅) が同一になることを確認する。
    /// </summary>
    [Fact]
    public void CreateRenderContextShouldKeepLayoutIndependentFromFontResolver()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
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

    /// <summary>
    /// <see cref="ReportRenderOption.PageSizeResolver"/> により任意ページサイズを指定できることを確認する。
    /// </summary>
    [Fact]
    public void BuildRenderPlanShouldUseConfiguredPageSizeResolver()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
        });

        var workbook = ExcelReader.Read(stream);
        var options = new ReportRenderOption
        {
            PageSizeResolver = static paperSize => paperSize == XLPaperSize.A4Paper ? (700d, 900d) : (595.28d, 841.89d)
        };

        var renderPlan = PdfRenderPlanner.BuildPlan(workbook, options);
        var page = Assert.Single(renderPlan[0].Pages);

        Assert.Equal(700d, page.PageBounds.Width, 3);
        Assert.Equal(900d, page.PageBounds.Height, 3);
    }

    /// <summary>
    /// <see cref="PdfRenderPlanner.BuildPlan"/> がセルの高さをコンテンツ領域に正しく反映することを確認する。
    /// </summary>
    [Fact]
    public void BuildRenderPlanShouldPreserveCellHeightForTextLayout()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
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

    /// <summary>
    /// <see cref="ReportDebugDumper"/> がワークブックおよび PDF 準備情報を JSON として出力することを確認する。
    /// </summary>
    [Fact]
    public void DebugDumperShouldWriteWorkbookAndPdfPreparationAsJson()
    {
        using var stream = TestWorkbookFactory.CreateWorkbook(workbook =>
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

    private static byte[] GetEmbeddedFontBytes() =>
        File.ReadAllBytes(TestHelper.IpaExGothicFontPath);

    private sealed class EmbeddedFontResolver : IReportFontResolver
    {
        private readonly ReadOnlyMemory<byte> fontData;

        public EmbeddedFontResolver(ReadOnlyMemory<byte> fontData)
        {
            this.fontData = fontData;
        }

        public FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic) =>
            String.Equals(familyName, "CustomEmbeddedFont", StringComparison.Ordinal)
                ? new FontResolveInfo("IPAexGothic")
                : null;

        public ReadOnlyMemory<byte>? GetFont(string faceName) =>
            String.Equals(faceName, "IPAexGothic", StringComparison.Ordinal)
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

        public FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic)
        {
            Requests.Add((familyName, bold, italic));
            return String.Equals(familyName, "CustomEmbeddedFont", StringComparison.Ordinal)
                ? new FontResolveInfo("IPAexGothic")
                : null;
        }

        public ReadOnlyMemory<byte>? GetFont(string faceName) =>
            String.Equals(faceName, "IPAexGothic", StringComparison.Ordinal)
                ? fontData
                : null;
    }
}
