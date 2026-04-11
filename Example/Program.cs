// <copyright file="Program.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

using System.Globalization;

using Example;

using OysterReport;

var inputPath = ResolveInputPath(args);
var (installedFontOutputPath, embeddedFontOutputPath) = ResolveOutputPaths(args, inputPath);

using var workbook = new TemplateWorkbook(inputPath);
FillInvoice(workbook.Sheets[0]);

var installedFontEngine = new OysterReportEngine
{
    FontResolver = new WindowsJapaneseFontResolver()
};
using (var output = File.Create(installedFontOutputPath))
{
    installedFontEngine.GeneratePdf(workbook, output);
}

var embeddedFontPath = Path.Combine(AppContext.BaseDirectory, "ipaexg.ttf");
var embeddedFontEngine = new OysterReportEngine
{
    FontResolver = new IpaExGothicFontResolver(embeddedFontPath)
};
using (var output = File.Create(embeddedFontOutputPath))
{
    embeddedFontEngine.GeneratePdf(workbook, output);
}

Console.WriteLine($"Input                : {inputPath}");
Console.WriteLine($"Installed font output: {installedFontOutputPath}");
Console.WriteLine($"Embedded font output : {embeddedFontOutputPath}");

// ---------------------------------------------------------------------------

static void FillInvoice(TemplateSheet sheet)
{
    // ---- Header fields ----
    sheet.ReplacePlaceholder("Title", "御請求書");
    sheet.ReplacePlaceholder("BillingTo", "株式会社サンプル");
    sheet.ReplacePlaceholder("IssueDate", "2025-04-11");
    sheet.ReplacePlaceholder("IssueNo", "INV-2025-001");
    sheet.ReplacePlaceholder("DeliveryDate", "2025-03-31");

    // ---- Detail rows (10 items) ----
    (int No, string Item, int Qty, int Price)[] items =
    [
        (1, "ソフトウェア開発", 1, 500_000),
        (2, "システム設計", 1, 200_000),
        (3, "テスト・品質保証", 2, 50_000),
        (4, "保守サービス", 12, 30_000),
        (5, "ライセンス費用", 5, 10_000),
        (6, "サーバー構築", 1, 80_000),
        (7, "トレーニング", 3, 40_000),
        (8, "ドキュメント作成", 1, 60_000),
        (9, "コンサルティング", 5, 50_000),
        (10, "緊急対応費", 1, 100_000),
    ];

    var templateRow = sheet.FindRow("No");
    var lastRow = templateRow;
    foreach (var (no, itemName, qty, price) in items)
    {
        lastRow = templateRow.InsertCopyAfter(lastRow);
        lastRow.ReplacePlaceholder("No", no.ToString(CultureInfo.InvariantCulture));
        lastRow.ReplacePlaceholder("Item", itemName);
        lastRow.ReplacePlaceholder("Qty", qty.ToString("N0", CultureInfo.InvariantCulture));
        lastRow.ReplacePlaceholder("Price", price.ToString("N0", CultureInfo.InvariantCulture));
        lastRow.ReplacePlaceholder("Amount", (qty * price).ToString("N0", CultureInfo.InvariantCulture));
    }

    templateRow.Delete();

    // ---- Totals ----
    var subTotal = items.Sum(static i => i.Qty * i.Price);
    var tax = (int)(subTotal * 0.1);
    var total = subTotal + tax;

    sheet.ReplacePlaceholder("SubTotal", subTotal.ToString("N0", CultureInfo.InvariantCulture));
    sheet.ReplacePlaceholder("Tax", tax.ToString("N0", CultureInfo.InvariantCulture));
    sheet.ReplacePlaceholder("TotalAmount", total.ToString("N0", CultureInfo.InvariantCulture));
}

static string ResolveInputPath(string[] args)
{
    if (args.Length > 0)
    {
        return Path.GetFullPath(args[0]);
    }

    var currentDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());
    while (currentDirectory is not null)
    {
        var candidate = Path.Combine(currentDirectory.FullName, "Invoice.xlsx");
        if (File.Exists(candidate))
        {
            return candidate;
        }

        var exampleCandidate = Path.Combine(currentDirectory.FullName, "Example", "Invoice.xlsx");
        if (File.Exists(exampleCandidate))
        {
            return exampleCandidate;
        }

        currentDirectory = currentDirectory.Parent;
    }

    throw new FileNotFoundException("Invoice.xlsx not found");
}

static (string InstalledFontOutputPath, string EmbeddedFontOutputPath) ResolveOutputPaths(string[] args, string inputPath)
{
    var outputDirectory = args.Length > 1
        ? Path.GetFullPath(args[1])
        : Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();

    return (
        Path.Combine(outputDirectory, "Invoice.installed-fonts.pdf"),
        Path.Combine(outputDirectory, "Invoice.ipaexg.pdf"));
}
