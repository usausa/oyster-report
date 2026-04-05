// <copyright file="Program.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

using OysterReport;
using OysterReport.Writing.Pdf;

var inputPath = ResolveInputPath(args);
var outputPath = ResolveOutputPath(args, inputPath);

var engine = new OysterReportEngine();
var workbook = engine.Read(inputPath);

// --- 罫線診断 ---
var sheet = workbook.Sheets[0];
foreach (var addr in new[] { "B1", "C1", "D1", "E1", "F1" })
{
    var cell = sheet.Cells.FirstOrDefault(c => c.Address == addr);
    if (cell is null)
    {
        Console.WriteLine($"{addr}: (cell not found)");
        continue;
    }

    var b = cell.Style.Borders;
    var f = cell.Style.Fill;
    Console.WriteLine($"{addr}: fill={f.BackgroundColorHex}  L={b.Left.Style} T={b.Top.Style} R={b.Right.Style} B={b.Bottom.Style}");
}

// --- 罫線診断終了 ---

var options = new PdfGenerateOptions
{
    FontResolver = new JapaneseFontResolver(),
};

using var output = File.Create(outputPath);
engine.GeneratePdf(workbook, output, options);

Console.WriteLine($"Input : {inputPath}");
Console.WriteLine($"Output: {outputPath}");

static string ResolveInputPath(string[] args)
{
    if (args.Length > 0)
    {
        return Path.GetFullPath(args[0]);
    }

    var currentDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());
    while (currentDirectory is not null)
    {
        var candidate = Path.Combine(currentDirectory.FullName, "seikyusyo.xlsx");
        if (File.Exists(candidate))
        {
            return candidate;
        }

        currentDirectory = currentDirectory.Parent;
    }

    throw new FileNotFoundException("seikyusyo.xlsx が見つかりません。引数で入力ファイルを指定してください。");
}

static string ResolveOutputPath(string[] args, string inputPath)
{
    if (args.Length > 1)
    {
        return Path.GetFullPath(args[1]);
    }

    return Path.ChangeExtension(inputPath, ".pdf");
}
