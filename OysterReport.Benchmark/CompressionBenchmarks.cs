namespace OysterReport.Benchmark;

using System.IO.Compression;

using BenchmarkDotNet.Attributes;

using ClosedXML.Excel.Drawings;

using OysterReport.Internal;
using OysterReport.Internal.OpenXml;

// Deflate level trade-off for a text/border-heavy invoice and an image-heavy sheet
[MemoryDiagnoser]
public class CompressionLevelBenchmark
{
    private const int InvoiceRows = 300;
    private const int Photos = 6;

    private ReportWorkbook invoice = default!;
    private ReportWorkbook photos = default!;
    private OysterReportEngine engine = default!;

    [Params(CompressionLevel.Optimal, CompressionLevel.Fastest, CompressionLevel.SmallestSize, CompressionLevel.NoCompression)]
    public CompressionLevel Level { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        engine = new OysterReportEngine { FontResolver = BenchmarkFontResolver.Instance, CompressionLevel = Level };

        invoice = Load(InvoiceTemplateFactory.Create(InvoiceRows));
        var sheet = new TemplateSheet(invoice, invoice.Sheets[0]);
        sheet.ReplacePlaceholders(InvoiceTemplateFactory.CreateHeader());
        sheet.ReplacePlaceholders(InvoiceTemplateFactory.CreateDetails(InvoiceRows));

        photos = Load(CreatePhotoWorkbook());
    }

    [Benchmark]
    public long Invoice() => Generate(invoice);

    [Benchmark]
    public long Photo() => Generate(photos);

    internal static byte[] CreatePhotoWorkbook() =>
        BenchmarkWorkbookFactory.Create(xl =>
        {
            var sheet = xl.AddWorksheet("Photos");

            // Keeps the image anchors inside the used range so the images are rendered
            sheet.Cell(60, 12).Value = "End";
            for (var i = 0; i < Photos; i++)
            {
                sheet.Cell(1 + ((i / 2) * 20), 1).Value = $"Photo{i + 1}";
                using var image = new MemoryStream(BenchmarkImageFactory.CreateRgbaPng(400, 300, (uint)(i + 1)), writable: false);
                sheet.AddPicture(image, XLPictureFormat.Png, $"Photo{i + 1}")
                    .MoveTo(sheet.Cell(2 + ((i / 2) * 20), 2 + ((i % 2) * 5)))
                    .WithSize(200, 150);
            }
        });

    private static ReportWorkbook Load(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return OpenXmlLoader.Load(stream);
    }

    private long Generate(ReportWorkbook workbook)
    {
        var sheet = new TemplateSheet(workbook, workbook.Sheets[0]);
        using var output = new MemoryStream();
        engine.GeneratePdf(sheet, output);
        return output.Length;
    }
}
