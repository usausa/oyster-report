// <copyright file="WorkbookTestFactory.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

using ClosedXML.Excel;

internal static class WorkbookTestFactory
{
    public static MemoryStream CreateWorkbook(Action<IXLWorkbook> configure)
    {
        using var workbook = new XLWorkbook();
        configure(workbook);

        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }
}
