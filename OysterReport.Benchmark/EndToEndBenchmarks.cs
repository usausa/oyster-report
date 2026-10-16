namespace OysterReport.Benchmark;

using System.Globalization;

using BenchmarkDotNet.Attributes;

using ClosedXML.Excel;

using OysterReport.Internal;
using OysterReport.Internal.OpenXml;

// Invoice-style flow with an embedded Japanese font: load template, clone and fill, generate PDF
[MemoryDiagnoser]
public class InvoiceEndToEndBenchmark
{
    private byte[] templateBytes = default!;
    private ReportWorkbook template = default!;
    private TemplateSheet filled = default!;
    private OysterReportEngine engine = default!;
    private Dictionary<string, string?> header = default!;
    private List<Dictionary<string, string?>> details = default!;

    [Params(30, 300)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        templateBytes = InvoiceTemplateFactory.Create(Rows);
        using (var stream = new MemoryStream(templateBytes, writable: false))
        {
            template = OpenXmlLoader.Load(stream);
        }

        engine = new OysterReportEngine { FontResolver = BenchmarkFontResolver.Instance };
        header = InvoiceTemplateFactory.CreateHeader();
        details = InvoiceTemplateFactory.CreateDetails(Rows);

        filled = CloneSheet();
        Fill(filled);
    }

    [Benchmark]
    public int Load()
    {
        using var stream = new MemoryStream(templateBytes, writable: false);
        using var workbook = new TemplateWorkbook(stream);
        return workbook.Sheets.Count;
    }

    [Benchmark]
    public int CloneAndFill() => Fill(CloneSheet());

    [Benchmark]
    public long GeneratePdf()
    {
        using var output = new MemoryStream();
        engine.GeneratePdf(filled, output);
        return output.Length;
    }

    [Benchmark]
    public long EndToEnd()
    {
        var sheet = CloneSheet();
        Fill(sheet);
        using var output = new MemoryStream();
        engine.GeneratePdf(sheet, output);
        return output.Length;
    }

    private TemplateSheet CloneSheet()
    {
        var workbook = template.DeepClone();
        return new TemplateSheet(workbook, workbook.Sheets[0]);
    }

    private int Fill(TemplateSheet sheet) =>
        sheet.ReplacePlaceholders(header) + sheet.ReplacePlaceholders(details);
}

// Detail expansion by copying a template row once per item (InsertCopyBelow)
[MemoryDiagnoser]
public class RowExpansionBenchmark
{
    private ReportWorkbook template = default!;

    [Params(30, 300)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        using var stream = new MemoryStream(InvoiceTemplateFactory.Create(1), writable: false);
        template = OpenXmlLoader.Load(stream);
    }

    [Benchmark]
    public int Expand()
    {
        var workbook = template.DeepClone();
        var sheet = new TemplateSheet(workbook, workbook.Sheets[0]);
        var row = sheet.FindRow("Item");
        for (var i = 1; i < Rows; i++)
        {
            row = row.InsertCopyBelow();
        }

        return row.RowNumber;
    }
}

internal static class InvoiceTemplateFactory
{
    private const int DetailStartRow = 7;

    private static readonly string[] ItemNames =
    [
        "CPU", "CPUクーラー", "マザーボード", "メモリ 32GB", "NVMe SSD 2TB",
        "グラフィックボード", "電源ユニット 850W", "PCケース", "ケースファン 140mm"
    ];

    public static byte[] Create(int rows) =>
        BenchmarkWorkbookFactory.Create(xl =>
        {
            xl.Style.Font.FontName = "游ゴシック";
            var sheet = xl.AddWorksheet("Invoice");
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
            sheet.PageSetup.Header.Center.AddText("{{InvoiceNo}}");
            sheet.PageSetup.Footer.Center.AddText("&P / &N");

            sheet.Cell("A1").Value = "{{Subject}}";
            sheet.Range("A1:F1").Merge();
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 16d;
            sheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            sheet.Cell("A3").Value = "請求先";
            sheet.Cell("B3").Value = "{{BillingTo}}";
            sheet.Cell("A4").Value = "請求日";
            sheet.Cell("B4").Value = "{{InvoiceDate}}";
            sheet.Cell("D3").Value = "請求番号";
            sheet.Cell("E3").Value = "{{InvoiceNo}}";
            sheet.Cell("D4").Value = "納品日";
            sheet.Cell("E4").Value = "{{DeliveryDate}}";

            string[] labels = ["No", "品目", "数量", "単価", "金額", "備考"];
            for (var c = 0; c < labels.Length; c++)
            {
                var cell = sheet.Cell(DetailStartRow - 1, c + 1);
                cell.Value = labels[c];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(220, 230, 241);
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            string[] markers = ["{{No}}", "{{Item}}", "{{Qty}}", "{{Price}}", "{{Amount}}", "{{Note}}"];
            for (var c = 0; c < markers.Length; c++)
            {
                sheet.Cell(DetailStartRow, c + 1).Value = markers[c];
            }

            var lastDetailRow = DetailStartRow + rows - 1;
            var table = sheet.Range(DetailStartRow - 1, 1, lastDetailRow, 6);
            table.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            sheet.Range(DetailStartRow, 3, lastDetailRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            for (var r = DetailStartRow + 1; r <= lastDetailRow; r += 2)
            {
                sheet.Range(r, 1, r, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242);
            }

            var totalRow = lastDetailRow + 2;
            sheet.Cell(totalRow, 4).Value = "合計";
            sheet.Cell(totalRow, 5).Value = "{{TotalAmount}}";
            sheet.Range(totalRow, 5, totalRow, 6).Merge();
            sheet.Range(totalRow, 4, totalRow, 6).Style.Border.TopBorder = XLBorderStyleValues.Double;
            sheet.Range(totalRow, 4, totalRow, 6).Style.Font.Underline = XLFontUnderlineValues.Single;

            sheet.PageSetup.PrintAreas.Add($"A1:F{totalRow}");
        });

    public static Dictionary<string, string?> CreateHeader() => new()
    {
        ["Subject"] = "御請求書",
        ["BillingTo"] = "株式会社サンプル",
        ["InvoiceDate"] = "2025-04-11",
        ["InvoiceNo"] = "INV-2025-001",
        ["DeliveryDate"] = "2025-03-31",
        ["TotalAmount"] = "1,234,567"
    };

    public static List<Dictionary<string, string?>> CreateDetails(int rows)
    {
        var result = new List<Dictionary<string, string?>>(rows);
        for (var i = 0; i < rows; i++)
        {
            var qty = (i % 3) + 1;
            var price = 1_980 + (i * 137);
            result.Add(new Dictionary<string, string?>
            {
                ["No"] = (i + 1).ToString(CultureInfo.InvariantCulture),
                ["Item"] = ItemNames[i % ItemNames.Length],
                ["Qty"] = qty.ToString("N0", CultureInfo.InvariantCulture),
                ["Price"] = price.ToString("N0", CultureInfo.InvariantCulture),
                ["Amount"] = (qty * price).ToString("N0", CultureInfo.InvariantCulture),
                ["Note"] = (i % 4) == 0 ? "特記事項あり" : null
            });
        }

        return result;
    }
}

internal sealed class BenchmarkFontResolver : IReportFontResolver
{
    private const string FaceName = "IPAexGothic";

    public static BenchmarkFontResolver Instance { get; } = new();

    private readonly ReadOnlyMemory<byte> fontData = File.ReadAllBytes(FindFont());

    public FontResolveInfo ResolveTypeface(string familyName, bool bold, bool italic) =>
        new(FaceName) { MustSimulateBold = bold, MustSimulateItalic = italic };

    public ReadOnlyMemory<byte>? GetFont(string faceName) =>
        String.Equals(faceName, FaceName, StringComparison.Ordinal) ? fontData : null;

    private static string FindFont()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var path = Path.Combine(directory.FullName, "ipaexg.ttf");
            if (File.Exists(path))
            {
                return path;
            }
        }

        throw new FileNotFoundException("ipaexg.ttf not found.");
    }
}
