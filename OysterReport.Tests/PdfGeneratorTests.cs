// <copyright file="PdfGeneratorTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

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

        public ReportFontResolveResult? ResolveFont(ReportFontRequest request)
        {
            if (!string.Equals(request.FontName, "CustomEmbeddedFont", StringComparison.Ordinal))
            {
                return null;
            }

            return new ReportFontResolveResult
            {
                FontName = "IPAexGothic",
                FontData = fontData
            };
        }
    }
}
