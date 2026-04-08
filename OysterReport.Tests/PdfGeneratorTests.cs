// <copyright file="PdfGeneratorTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using OysterReport.Generator;

using Xunit;

public sealed class PdfGeneratorTests
{
    [Fact]
    public void GenerateShouldCreatePdfDocument()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
            sheet.Cell("A2").Value = "World";
        });

        var workbook = new ExcelReader().Read(stream);
        using var output = new MemoryStream();

        new PdfGenerator().Generate(workbook, output, new PdfGeneratorOption());

        output.Position = 0;
        using var reader = new StreamReader(output, leaveOpen: true);
        var header = reader.ReadLine();
        Assert.NotNull(header);
        Assert.StartsWith("%PDF", header, StringComparison.Ordinal);
    }

    [Fact]
    public void DebugDumperShouldWriteWorkbookAndPdfPreparation()
    {
        using var stream = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "Hello";
        });

        var workbook = new ExcelReader().Read(stream);
        var dumper = new ReportDebugDumper();

        using var workbookDump = new MemoryStream();
        dumper.DumpWorkbook(workbook, workbookDump);
        var workbookJson = System.Text.Encoding.UTF8.GetString(workbookDump.ToArray());
        Assert.Contains("\"Sheets\"", workbookJson, StringComparison.Ordinal);

        using var pdfPreparationDump = new MemoryStream();
        dumper.DumpPdfPreparation(workbook, pdfPreparationDump);
        var preparationJson = System.Text.Encoding.UTF8.GetString(pdfPreparationDump.ToArray());
        Assert.Contains("\"RenderPlan\"", preparationJson, StringComparison.Ordinal);
    }
}
