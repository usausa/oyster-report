namespace Example;

using System.Globalization;

using OysterReport;

using Smart.CommandLine.Hosting;

//--------------------------------------------------------------------------------
// Command builder
//--------------------------------------------------------------------------------
public static class CommandBuilderExtensions
{
    public static void AddCommands(this ICommandBuilder commands)
    {
        commands.AddCommand<InvoiceCommand>();
        commands.AddCommand<EmbeddedCommand>();
    }
}

//--------------------------------------------------------------------------------
// invoice
//--------------------------------------------------------------------------------
[Command("invoice", "Place holder example")]
public sealed class InvoiceCommand : ICommandHandler
{
    private static readonly (int No, string Item, int Qty, int Price)[] Items =
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

    public ValueTask ExecuteAsync(CommandContext context)
    {
        var engine = new OysterReportEngine
        {
            FontResolver = new WindowsJapaneseFontResolver()
        };

        using var workbook = new TemplateWorkbook("Invoice.xlsx");
        EditSheet(workbook.Sheets[0]);

        using var output = File.Create("Invoice.pdf");
        engine.GeneratePdf(workbook, output);

        return ValueTask.CompletedTask;
    }

    private static void EditSheet(TemplateSheet sheet)
    {
        sheet.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["Subject"] = "御請求書",
            ["BillingTo"] = "株式会社サンプル",
            ["InvoiceDate"] = "2025-04-11",
            ["InvoiceNo"] = "INV-2025-001",
            ["DeliveryDate"] = "2025-03-31"
        });

        sheet.ReplacePlaceholders(Items.Select(static i => new Dictionary<string, string?>
        {
            ["No"] = i.No.ToString(CultureInfo.InvariantCulture),
            ["Item"] = i.Item,
            ["Qty"] = i.Qty.ToString("N0", CultureInfo.InvariantCulture),
            ["Price"] = i.Price.ToString("N0", CultureInfo.InvariantCulture),
            ["Amount"] = (i.Qty * i.Price).ToString("N0", CultureInfo.InvariantCulture)
        }));

        var subTotal = Items.Sum(static i => i.Qty * i.Price);
        var tax = (int)(subTotal * 0.1);
        sheet.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["SubTotal"] = subTotal.ToString("N0", CultureInfo.InvariantCulture),
            ["Tax"] = tax.ToString("N0", CultureInfo.InvariantCulture),
            ["TotalAmount"] = (subTotal + tax).ToString("N0", CultureInfo.InvariantCulture)
        });
    }
}

//--------------------------------------------------------------------------------
// embedded
//--------------------------------------------------------------------------------
[Command("embedded", "Embedded font example")]
public sealed class EmbeddedCommand : ICommandHandler
{
    public ValueTask ExecuteAsync(CommandContext context)
    {
        using var workbook = new TemplateWorkbook("Embedded.xlsx");

        var engine = new OysterReportEngine
        {
            FontResolver = new EmbeddedFontResolver()
        };

        using var output = File.Create("embedded.pdf");
        engine.GeneratePdf(workbook, output);

        return ValueTask.CompletedTask;
    }
}
