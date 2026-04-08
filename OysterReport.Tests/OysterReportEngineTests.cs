// <copyright file="OysterReportEngineTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

using Xunit;

public sealed class OysterReportEngineTests
{
    [Fact]
    public void EngineShouldSupportEndToEndFlow()
    {
        using var input = WorkbookTestFactory.CreateWorkbook(workbook =>
        {
            var sheet = workbook.AddWorksheet("Report");
            sheet.Cell("A1").Value = "{{Name}}";
        });

        var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (var file = File.Create(tempFile))
            {
                input.CopyTo(file);
            }

            var engine = new OysterReportEngine();
            using var workbook = new TemplateWorkbook(new XLWorkbook(tempFile));
            var sheet = Assert.Single(workbook.Sheets);
            sheet.ReplacePlaceholder("Name", "Bob");

            using var output = new MemoryStream();
            engine.GeneratePdf(workbook, output);

            Assert.True(output.Length > 0);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
