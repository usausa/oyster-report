namespace OysterReport.Benchmark;

using BenchmarkDotNet.Attributes;

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

using OysterReport.Internal;
using OysterReport.Internal.OpenXml;

// Loading a sparse sheet whose print area is far larger than the actual data
[MemoryDiagnoser]
public class SparseSheetLoadBenchmark
{
    private byte[] templateBytes = default!;

    [Params(1000, 10000)]
    public int PrintAreaRows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        templateBytes = BenchmarkWorkbookFactory.Create(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Header";
            sheet.Cell("B2").Value = 123;
            sheet.PageSetup.PrintAreas.Add($"A1:Z{PrintAreaRows}");
        });
    }

    [Benchmark]
    public int Load()
    {
        using var stream = new MemoryStream(templateBytes, writable: false);
        using var workbook = new TemplateWorkbook(stream);
        return workbook.ReportWorkbook.Sheets[0].Cells.Count;
    }
}

// Multi-key placeholder replacement over many detail rows
[MemoryDiagnoser]
public class ReplacePlaceholdersBenchmark
{
    private const int KeysPerRow = 8;

    private ReportWorkbook workbook = default!;
    private Dictionary<string, string?> values = default!;

    [Params(100, 1000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var bytes = BenchmarkWorkbookFactory.Create(xl =>
        {
            var sheet = xl.AddWorksheet("Report");
            for (var r = 1; r <= Rows; r++)
            {
                for (var c = 1; c <= KeysPerRow; c++)
                {
                    sheet.Cell(r, c).Value = $"{{{{Key{c}}}}}";
                }
            }
        });

        using var stream = new MemoryStream(bytes, writable: false);
        workbook = OpenXmlLoader.Load(stream);

        values = [];
        for (var c = 1; c <= KeysPerRow; c++)
        {
            values[$"Key{c}"] = $"Value{c}";
        }
    }

    // Includes DeepClone to model the clone-then-fill flow used per report
    [Benchmark]
    public int Replace()
    {
        var clone = workbook.DeepClone();
        var sheet = new TemplateSheet(clone, clone.Sheets[0]);
        return sheet.ReplacePlaceholders(values);
    }
}

// Render plan generation for a sheet with merged ranges and a striped table
[MemoryDiagnoser]
public class RenderPlanBenchmark
{
    private ReportWorkbook workbook = default!;

    [Params(500, 2000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var bytes = BenchmarkWorkbookFactory.Create(xl =>
        {
            var sheet = xl.AddWorksheet("Report");
            for (var r = 1; r <= Rows; r++)
            {
                for (var c = 1; c <= 6; c++)
                {
                    sheet.Cell(r, c).Value = $"R{r}C{c}";
                }
            }

            var table = sheet.Range($"A1:F{Rows}").CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2;

            for (var r = 10; r <= Rows; r += 50)
            {
                sheet.Range($"G{r}:I{r}").Merge();
                sheet.Cell($"G{r}").Value = "Merged";
            }
        });

        using var stream = new MemoryStream(bytes, writable: false);
        workbook = OpenXmlLoader.Load(stream);
    }

    [Benchmark]
    public int BuildPlan() => PdfRenderPlanner.BuildPlan(workbook).Count;
}

// End-to-end PDF generation for a sheet containing images
[MemoryDiagnoser]
public class PdfGenerationBenchmark
{
    private static readonly byte[] OnePxPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kZs8AAAAASUVORK5CYII=");

    private ReportWorkbook workbook = default!;
    private OysterReportEngine engine = default!;

    [Params(20)]
    public int Images { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var bytes = BenchmarkWorkbookFactory.Create(xl =>
        {
            var sheet = xl.AddWorksheet("Report");

            // Keeps the image anchor column inside the used range so the images are rendered
            sheet.Cell(1, 3).Value = "End";
            for (var r = 1; r <= Images; r++)
            {
                sheet.Cell(r, 1).Value = $"Row{r}";
                using var image = new MemoryStream(OnePxPng, writable: false);
                sheet.AddPicture(image, XLPictureFormat.Png, $"Image{r}")
                    .MoveTo(sheet.Cell(r, 2))
                    .WithSize(40, 12);
            }
        });

        using var stream = new MemoryStream(bytes, writable: false);
        workbook = OpenXmlLoader.Load(stream);
        engine = new OysterReportEngine();
    }

    [Benchmark]
    public long GeneratePdf()
    {
        var sheet = new TemplateSheet(workbook, workbook.Sheets[0]);
        using var output = new MemoryStream();
        engine.GeneratePdf(sheet, output);
        return output.Length;
    }
}

internal static class BenchmarkWorkbookFactory
{
    public static byte[] Create(Action<IXLWorkbook> configure)
    {
        using var workbook = new XLWorkbook();
        configure(workbook);
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }
}
