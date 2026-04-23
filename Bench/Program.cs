// OysterReport Invoice.xlsx benchmark — measures Load / Edit / GeneratePdf separately.

using System.Diagnostics;
using System.Globalization;
using System.Reflection;

using OysterReport;

const string XlsxPath = "Invoice.xlsx";
const int Warmup = 5;
const int Iterations = 30;

if (!File.Exists(XlsxPath))
{
    Console.WriteLine($"Invoice.xlsx not found at {Path.GetFullPath(XlsxPath)}");
    return 1;
}

// Page cache the file
_ = File.ReadAllBytes(XlsxPath);

Console.WriteLine($"File: {XlsxPath}  ({new FileInfo(XlsxPath).Length:N0} bytes)");
Console.WriteLine($"Warmup={Warmup}  Iterations={Iterations}");
Console.WriteLine();

// JIT warmup
for (var i = 0; i < Warmup; i++)
{
    _ = RunOnce();
}

var loadTicks = new long[Iterations];
var editTicks = new long[Iterations];
var genTicks = new long[Iterations];
var readTicks = new long[Iterations];
var pdfTicks = new long[Iterations];
var totalTicks = new long[Iterations];
var pdfBytes = 0L;

var createRenderContext = typeof(OysterReportEngine).GetMethod(
    "CreateRenderContext",
    BindingFlags.Instance | BindingFlags.NonPublic,
    types: [typeof(TemplateWorkbook)]) ?? throw new InvalidOperationException("CreateRenderContext(TemplateWorkbook) not found");
var pdfGeneratorType = typeof(OysterReportEngine).Assembly.GetType("OysterReport.Internal.PdfGenerator")
    ?? throw new InvalidOperationException("PdfGenerator type not found");
var writePdf = pdfGeneratorType.GetMethod(
    "WritePdf",
    BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public)
    ?? throw new InvalidOperationException("WritePdf method not found");

GC.Collect();
GC.WaitForPendingFinalizers();
GC.Collect();
var allocStart = GC.GetTotalAllocatedBytes(precise: true);

for (var i = 0; i < Iterations; i++)
{
    var (load, edit, gen, size) = RunOnce();
    loadTicks[i] = load;
    editTicks[i] = edit;
    genTicks[i] = gen;
    totalTicks[i] = load + edit + gen;
    pdfBytes = size;
}

// Drill-down: split Gen into ExcelReader (build context) vs PDF emission
for (var i = 0; i < Iterations; i++)
{
    var (read, pdf) = RunGenSplit(createRenderContext, writePdf);
    readTicks[i] = read;
    pdfTicks[i] = pdf;
}

var allocEnd = GC.GetTotalAllocatedBytes(precise: true);

Report("Load   (new TemplateWorkbook)", loadTicks);
Report("Edit   (placeholder + row ops)", editTicks);
Report("Gen    (GeneratePdf = ExcelReader + PDF)", genTicks);
Report("  Read   (ExcelReader → ReportWorkbook)", readTicks);
Report("  PDF    (PdfGenerator.WritePdf)", pdfTicks);
Report("Total", totalTicks);

Console.WriteLine();
Console.WriteLine($"PDF size: {pdfBytes:N0} bytes");
Console.WriteLine($"Allocated across {Iterations} runs: {(allocEnd - allocStart) / 1024.0 / 1024.0:F2} MB  " +
                  $"(per-run avg: {(allocEnd - allocStart) / (double)Iterations / 1024.0:F0} KB)");
Console.WriteLine();

var loadShare = Share(loadTicks, totalTicks);
var editShare = Share(editTicks, totalTicks);
var genShare = Share(genTicks, totalTicks);
Console.WriteLine($"Share of total (median basis):  Load={loadShare:P1}  Edit={editShare:P1}  Gen={genShare:P1}");
return 0;

static (long Load, long Edit, long Gen, long PdfLength) RunOnce()
{
    var sw = new Stopwatch();

    sw.Restart();
    using var workbook = new TemplateWorkbook(XlsxPath);
    sw.Stop();
    var load = sw.ElapsedTicks;

    sw.Restart();
    EditSheet(workbook.Sheets[0]);
    sw.Stop();
    var edit = sw.ElapsedTicks;

    var engine = new OysterReportEngine { FontResolver = new JapaneseFontResolver() };
    using var output = new MemoryStream(capacity: 128 * 1024);
    sw.Restart();
    engine.GeneratePdf(workbook, output);
    sw.Stop();
    var gen = sw.ElapsedTicks;

    return (load, edit, gen, output.Length);
}

static (long Read, long Pdf) RunGenSplit(MethodInfo createCtx, MethodInfo writePdf)
{
    using var workbook = new TemplateWorkbook(XlsxPath);
    EditSheet(workbook.Sheets[0]);

    var engine = new OysterReportEngine { FontResolver = new JapaneseFontResolver() };
    var sw = new Stopwatch();

    sw.Restart();
    var context = createCtx.Invoke(engine, [workbook])!;
    sw.Stop();
    var read = sw.ElapsedTicks;

    using var output = new MemoryStream(capacity: 128 * 1024);
    sw.Restart();
    writePdf.Invoke(null, [context, output]);
    sw.Stop();
    var pdf = sw.ElapsedTicks;

    return (read, pdf);
}

static void EditSheet(TemplateSheet sheet)
{
    sheet.ReplacePlaceholder("Subject", "御請求書");
    sheet.ReplacePlaceholder("BillingTo", "株式会社サンプル");
    sheet.ReplacePlaceholder("InvoiceDate", "2025-04-11");
    sheet.ReplacePlaceholder("InvoiceNo", "INV-2025-001");
    sheet.ReplacePlaceholder("DeliveryDate", "2025-03-31");

    (int No, string Item, int Qty, int Price)[] items =
    [
        (1, "CPU", 1, 54_800),
        (2, "CPUクーラー", 1, 8_980),
        (3, "マザーボード", 1, 24_800),
        (4, "メモリ 32GB", 2, 15_600),
        (5, "NVMe SSD 2TB", 1, 18_400),
        (6, "グラフィックボード", 1, 89_800),
        (7, "電源ユニット 850W", 1, 16_200),
        (8, "PCケース", 1, 12_500),
        (9, "ケースファン 140mm", 3, 1_980)
    ];

    var templateRow = sheet.FindRow("No");
    var bottomRowNumber = sheet.FindRow("SubTotal").RowNumber - 1;
    var row = templateRow;
    foreach (var (no, itemName, qty, price) in items)
    {
        row = templateRow.InsertCopyAfter(row);
        row.ReplacePlaceholder("No", no.ToString(CultureInfo.InvariantCulture));
        row.ReplacePlaceholder("Item", itemName);
        row.ReplacePlaceholder("Qty", qty.ToString("N0", CultureInfo.InvariantCulture));
        row.ReplacePlaceholder("Price", price.ToString("N0", CultureInfo.InvariantCulture));
        row.ReplacePlaceholder("Amount", (qty * price).ToString("N0", CultureInfo.InvariantCulture));
        sheet.DeleteRows(bottomRowNumber, bottomRowNumber);
    }

    templateRow.Delete();

    var subTotal = items.Sum(static i => i.Qty * i.Price);
    var tax = (int)(subTotal * 0.1);
    var total = subTotal + tax;

    sheet.ReplacePlaceholder("SubTotal", subTotal.ToString("N0", CultureInfo.InvariantCulture));
    sheet.ReplacePlaceholder("Tax", tax.ToString("N0", CultureInfo.InvariantCulture));
    sheet.ReplacePlaceholder("TotalAmount", total.ToString("N0", CultureInfo.InvariantCulture));
}

static void Report(string name, long[] ticks)
{
    var ms = ticks.Select(t => t * 1000.0 / Stopwatch.Frequency).OrderBy(v => v).ToArray();
    var min = ms[0];
    var p50 = ms[ms.Length / 2];
    var p95 = ms[(int)(ms.Length * 0.95)];
    var max = ms[^1];
    var mean = ms.Average();
    Console.WriteLine($"{name,-42}  min={min,7:F2}  p50={p50,7:F2}  mean={mean,7:F2}  p95={p95,7:F2}  max={max,7:F2}   ms");
}

static double Share(long[] part, long[] whole)
{
    var p = part.Select(t => t * 1000.0 / Stopwatch.Frequency).OrderBy(v => v).ToArray();
    var w = whole.Select(t => t * 1000.0 / Stopwatch.Frequency).OrderBy(v => v).ToArray();
    return p[p.Length / 2] / w[w.Length / 2];
}

internal sealed class JapaneseFontResolver : IReportFontResolver
{
    private static readonly Dictionary<string, string> FontMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ＭＳ Ｐゴシック"] = "MS PGothic",
        ["MS Pゴシック"] = "MS PGothic",
        ["ＭＳ ゴシック"] = "MS Gothic",
        ["ＭＳ Ｐ明朝"] = "MS PMincho",
        ["MS P明朝"] = "MS PMincho",
        ["ＭＳ 明朝"] = "MS Mincho"
    };

    public FontResolveInfo? ResolveTypeface(string familyName, bool bold, bool italic) =>
        FontMap.TryGetValue(familyName, out var resolved) ? new FontResolveInfo(resolved) : null;
}
