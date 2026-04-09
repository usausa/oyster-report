// <copyright file="Program.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

using Example;

using OysterReport;

var inputPath = ResolveInputPath(args);
var (installedFontOutputPath, embeddedFontOutputPath) = ResolveOutputPaths(args, inputPath);

using var workbook = new TemplateWorkbook(inputPath);
workbook.ReplacePlaceholder("TotalAmount", "123,456,789");

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

Console.WriteLine($"Input : {inputPath}");
Console.WriteLine($"Installed font output: {installedFontOutputPath}");
Console.WriteLine($"Embedded font output : {embeddedFontOutputPath}");

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

        var exampleCandidate = Path.Combine(currentDirectory.FullName, "Example", "seikyusyo.xlsx");
        if (File.Exists(exampleCandidate))
        {
            return exampleCandidate;
        }

        currentDirectory = currentDirectory.Parent;
    }

    throw new FileNotFoundException("seikyusyo.xlsx not found");
}

static (string InstalledFontOutputPath, string EmbeddedFontOutputPath) ResolveOutputPaths(string[] args, string inputPath)
{
    var baseOutputPath = Path.ChangeExtension(inputPath, ".pdf");
    if (args.Length > 1)
    {
        baseOutputPath = Path.GetFullPath(args[1]);
    }

    var outputDirectory = Path.GetDirectoryName(baseOutputPath) ?? Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
    var outputFileNameWithoutExtension = Path.GetFileNameWithoutExtension(baseOutputPath);

    return (
        Path.Combine(outputDirectory, outputFileNameWithoutExtension + ".installed-fonts.pdf"),
        Path.Combine(outputDirectory, outputFileNameWithoutExtension + ".ipaexg.pdf"));
}
