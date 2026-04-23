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
    public ValueTask ExecuteAsync(CommandContext context)
    {
        using var workbook = new TemplateWorkbook("Invoice.xlsx");
        EditSheet(workbook.Sheets[0]);

        var engine = new OysterReportEngine
        {
            FontResolver = new WindowsJapaneseFontResolver()
        };

        using var output = File.Create("Invoice.pdf");
        engine.GeneratePdf(workbook, output);

        return ValueTask.CompletedTask;
    }

    private static void EditSheet(TemplateSheet sheet)
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
}

//--------------------------------------------------------------------------------
// embedded
//--------------------------------------------------------------------------------
[Command("embedded", "Embedded font example")]
public sealed class EmbeddedCommand : ICommandHandler
{
    public ValueTask ExecuteAsync(CommandContext context)
    {
        using var workbook = new TemplateWorkbook("Sample.xlsx");

        var engine = new OysterReportEngine
        {
            FontResolver = new EmbeddedFontResolver()
        };

        using var output = File.Create("embedded.pdf");
        engine.GeneratePdf(workbook, output);

        return ValueTask.CompletedTask;
    }
}
