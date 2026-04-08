// <copyright file="Program.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

using ClosedXML.Excel;

using Example;

using OysterReport;

var inputPath = ResolveInputPath(args);
var outputPath = ResolveOutputPath(args, inputPath);

var engine = new OysterReportEngine
{
    FontResolver = new JapaneseFontResolver()
};

using var workbook = new TemplateWorkbook(new XLWorkbook(inputPath));

using var output = File.Create(outputPath);
engine.GeneratePdf(workbook, output);

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

    throw new FileNotFoundException("seikyusyo.xlsx not found");
}

static string ResolveOutputPath(string[] args, string inputPath)
{
    if (args.Length > 1)
    {
        return Path.GetFullPath(args[1]);
    }

    return Path.ChangeExtension(inputPath, ".pdf");
}
