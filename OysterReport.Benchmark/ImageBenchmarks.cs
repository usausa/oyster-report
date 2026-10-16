namespace OysterReport.Benchmark;

using BenchmarkDotNet.Attributes;

using ClosedXML.Excel.Drawings;

using OysterReport.Internal;
using OysterReport.Internal.OpenXml;

// The same logo placed repeatedly on one sheet and once on each of several sheets
[MemoryDiagnoser]
public class ImageDedupBenchmark
{
    private const int Sheets = 5;

    private ReportWorkbook sameSheet = default!;
    private ReportWorkbook multiSheet = default!;
    private OysterReportEngine engine = default!;

    [Params(1, 10)]
    public int Placements { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        engine = new OysterReportEngine { FontResolver = BenchmarkFontResolver.Instance };
        var logo = BenchmarkImageFactory.CreateRgbaPng(300, 120, 7);
        sameSheet = Load(CreateWorkbook(logo, 1, Placements));
        multiSheet = Load(CreateWorkbook(logo, Sheets, Placements));
    }

    [Benchmark]
    public long SameSheet() => Generate(sameSheet);

    [Benchmark]
    public long MultiSheet() => Generate(multiSheet);

    internal static byte[] CreateWorkbook(byte[] logo, int sheets, int placements) =>
        BenchmarkWorkbookFactory.Create(xl =>
        {
            for (var s = 1; s <= sheets; s++)
            {
                var sheet = xl.AddWorksheet($"Sheet{s}");

                // Keeps the image anchors inside the used range so the images are rendered
                sheet.Cell(placements * 4, 8).Value = "End";
                for (var i = 0; i < placements; i++)
                {
                    sheet.Cell((i * 4) + 1, 1).Value = $"Row{i + 1}";
                    using var image = new MemoryStream(logo, writable: false);
                    sheet.AddPicture(image, XLPictureFormat.Png, $"Logo{s}_{i + 1}")
                        .MoveTo(sheet.Cell((i * 4) + 1, 3))
                        .WithSize(150, 60);
                }
            }
        });

    private static ReportWorkbook Load(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return OpenXmlLoader.Load(stream);
    }

    // Same steps as OysterReportEngine.GeneratePdf(TemplateWorkbook, Stream), without reloading the template
    private long Generate(ReportWorkbook workbook)
    {
        var context = new ReportRenderContext
        {
            Workbook = workbook,
            SheetPlans = PdfRenderPlanner.BuildPlan(workbook, engine.RenderingOptions),
            FontResolver = engine.FontResolver,
            RenderingOptions = engine.RenderingOptions
        };
        using var output = new MemoryStream();
        PdfGenerator.WritePdf(context, output);
        return output.Length;
    }
}
