// Prototype harness: runs the OpenXML-only loader against Invoice.xlsx, writes a PDF via the
// existing PdfGenerator (bypassing ClosedXML + ExcelReader), then benchmarks vs. current flow.

using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;

using OysterReport;
using OysterReport.Internal;
using OysterReport.Prototype;

const string XlsxPath = "Invoice.xlsx";
const int Warmup = 5;
const int Iterations = 30;

if (!File.Exists(XlsxPath))
{
    Console.WriteLine($"Invoice.xlsx not found at {Path.GetFullPath(XlsxPath)}");
    return 1;
}

_ = File.ReadAllBytes(XlsxPath);
Console.WriteLine($"File: {XlsxPath}  ({new FileInfo(XlsxPath).Length:N0} bytes)");
Console.WriteLine();

// Sanity smoke test — write one PDF from each path and print diagnostics.
var fontResolver = new JapaneseFontResolver();
var pdfGeneratorType = typeof(OysterReportEngine).Assembly.GetType("OysterReport.Internal.PdfGenerator")
    ?? throw new InvalidOperationException("PdfGenerator type not found");
var writePdf = pdfGeneratorType.GetMethod("WritePdf", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public)
    ?? throw new InvalidOperationException("WritePdf method not found");
var pdfRenderPlannerType = typeof(OysterReportEngine).Assembly.GetType("OysterReport.Internal.PdfRenderPlanner")
    ?? throw new InvalidOperationException("PdfRenderPlanner type not found");
var buildPlan = pdfRenderPlannerType.GetMethod("BuildPlan", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public)
    ?? throw new InvalidOperationException("BuildPlan method not found");
var renderContextType = typeof(OysterReportEngine).Assembly.GetType("OysterReport.Internal.ReportRenderContext")
    ?? throw new InvalidOperationException("ReportRenderContext type not found");

try
{
    using var stream = File.OpenRead(XlsxPath);
    var wb = OpenXmlLoader.Load(stream);
    Console.WriteLine($"[Prototype] Loaded workbook: sheets={wb.Sheets.Count}");
    foreach (var sheet in wb.Sheets)
    {
        Console.WriteLine($"  Sheet '{sheet.Name}': Used={sheet.UsedRange}  Cells={sheet.Cells.Count}  " +
                          $"Rows={sheet.Rows.Count}  Cols={sheet.Columns.Count}  Merges={sheet.MergedRanges.Count}  Images={sheet.Images.Count}");
    }

    // Write prototype PDF.
    using var output = File.Create("Invoice.prototype.pdf");
    WritePdfFromWorkbook(wb, output, fontResolver, buildPlan, writePdf, renderContextType);
    Console.WriteLine($"[Prototype] Wrote Invoice.prototype.pdf  ({output.Length:N0} bytes)");
}
catch (IOException ex)
{
    Console.WriteLine($"[Prototype] FAILED: {ex.GetType().Name}: {ex.Message}");
    return 2;
}
catch (InvalidOperationException ex)
{
    Console.WriteLine($"[Prototype] FAILED: {ex.GetType().Name}: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
    return 2;
}

// Sanity: also write the current-flow PDF for visual diff.
{
    using var tpl = new TemplateWorkbook(XlsxPath);
    var engine = new OysterReportEngine { FontResolver = fontResolver };
    using var output = File.Create("Invoice.current.pdf");
    engine.GeneratePdf(tpl, output);
    Console.WriteLine($"[Current]   Wrote Invoice.current.pdf    ({output.Length:N0} bytes)");
}

Console.WriteLine();

// ---------- Rigorous equivalence check ----------
ReportWorkbook currentWorkbook;
ReportWorkbook prototypeWorkbook;
{
    using var tpl = new TemplateWorkbook(XlsxPath);
    currentWorkbook = ExcelReader.Read(tpl.UnderlyingWorkbook);
}
{
    using var stream = File.OpenRead(XlsxPath);
    prototypeWorkbook = OpenXmlLoader.Load(stream);
}

Console.WriteLine($"=== ReportWorkbook deep-diff (Current vs Prototype) for {Path.GetFileName(XlsxPath)} ===");

var diffs = WorkbookDiff.Compare(currentWorkbook, prototypeWorkbook);
Console.WriteLine(WorkbookDiff.Summarize(diffs));

var currentPdfHash = HashFile("Invoice.current.pdf");
var prototypePdfHash = HashFile("Invoice.prototype.pdf");
Console.WriteLine($"=== PDF byte-level comparison ({Path.GetFileName(XlsxPath)}) ===");
Console.WriteLine($"  Invoice.current.pdf    SHA-256={currentPdfHash}");
Console.WriteLine($"  Invoice.prototype.pdf  SHA-256={prototypePdfHash}");
var pdfVerdict = currentPdfHash == prototypePdfHash
    ? "BYTE-IDENTICAL"
    : "DIFFER (even with equal workbooks, PDF metadata like CreationDate can differ)";
Console.WriteLine($"  Verdict: {pdfVerdict} (hash prefix {currentPdfHash[..8]} vs {prototypePdfHash[..8]})");

Console.WriteLine();

// ---------- Benchmarks ----------
for (var i = 0; i < Warmup; i++)
{
    _ = LoadCurrent();
    _ = LoadPrototype();
}

var currentLoad = new long[Iterations];
var currentRead = new long[Iterations];
var protoLoad = new long[Iterations];

for (var i = 0; i < Iterations; i++)
{
    var (load, read) = LoadCurrent();
    currentLoad[i] = load;
    currentRead[i] = read;
}
for (var i = 0; i < Iterations; i++)
{
    protoLoad[i] = LoadPrototype();
}

Console.WriteLine($"Warmup={Warmup}  Iterations={Iterations}");
var header = $"Measuring xlsx-load -> ReportWorkbook (the path being replaced) [iters={Iterations}]";
Console.WriteLine(header);
Report("  Current Load   (ClosedXML XLWorkbook)", currentLoad);
Report("  Current Read   (ExcelReader → ReportWorkbook)", currentRead);
Report("  Current TOTAL  (Load + Read)", currentLoad.Zip(currentRead, static (a, b) => a + b).ToArray());
Report("  Prototype     (OpenXML → ReportWorkbook)", protoLoad);

return 0;

(long Load, long Read) LoadCurrent()
{
    var sw = new Stopwatch();
    sw.Restart();
    using var wb = new TemplateWorkbook(XlsxPath);
    sw.Stop();
    var load = sw.ElapsedTicks;

    sw.Restart();
    _ = ExcelReader.Read(wb.UnderlyingWorkbook);
    sw.Stop();
    return (load, sw.ElapsedTicks);
}

long LoadPrototype()
{
    var sw = new Stopwatch();
    sw.Restart();
    using var stream = File.OpenRead(XlsxPath);
    _ = OpenXmlLoader.Load(stream);
    sw.Stop();
    return sw.ElapsedTicks;
}

static string HashFile(string path)
{
    using var stream = File.OpenRead(path);
    var hash = SHA256.HashData(stream);
    return Convert.ToHexString(hash);
}

static void Report(string name, long[] ticks)
{
    var ms = ticks.Select(t => t * 1000.0 / Stopwatch.Frequency).OrderBy(v => v).ToArray();
    var min = ms[0];
    var p50 = ms[ms.Length / 2];
    var p95 = ms[(int)(ms.Length * 0.95)];
    var max = ms[^1];
    var mean = ms.Average();
    Console.WriteLine($"{name,-48}  min={min,7:F2}  p50={p50,7:F2}  mean={mean,7:F2}  p95={p95,7:F2}  max={max,7:F2}   ms");
}

static void WritePdfFromWorkbook(
    ReportWorkbook workbook,
    Stream output,
    IReportFontResolver fontResolver,
    MethodInfo buildPlan,
    MethodInfo writePdf,
    Type renderContextType)
{
    var renderOption = new ReportRenderOption();
    var plans = buildPlan.Invoke(null, [workbook, renderOption])!;
    var context = Activator.CreateInstance(renderContextType)!;

    renderContextType.GetProperty("Workbook")!.SetValue(context, workbook);
    renderContextType.GetProperty("SheetPlans")!.SetValue(context, plans);
    renderContextType.GetProperty("FontResolver")!.SetValue(context, fontResolver);
    renderContextType.GetProperty("RenderingOptions")!.SetValue(context, renderOption);
    renderContextType.GetProperty("EmbedDocumentMetadata")!.SetValue(context, true);
    renderContextType.GetProperty("CompressContentStreams")!.SetValue(context, true);

    writePdf.Invoke(null, [context, output]);
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
