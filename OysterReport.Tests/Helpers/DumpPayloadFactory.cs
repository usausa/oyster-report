// <copyright file="DumpPayloadFactory.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests.Helpers;

using System.Globalization;
using System.Text.Json;

using OysterReport.Internal;

internal static class DumpPayloadFactory
{
    public static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true
    };

    public static object CreateWorkbookPayload(ReportWorkbook workbook) =>
        new
        {
            workbook.Metadata,
            workbook.MeasurementProfile,
            Sheets = workbook.Sheets.Select(sheet => new
            {
                sheet.Name,
                UsedRange = sheet.UsedRange.ToString(),
                sheet.ShowGridLines,
                Rows = sheet.Rows.Select(row => new
                {
                    row.Index,
                    row.HeightPoint,
                    row.TopPoint,
                    row.IsHidden,
                    row.OutlineLevel
                }),
                Columns = sheet.Columns.Select(column => new
                {
                    column.Index,
                    column.WidthPoint,
                    column.LeftPoint,
                    column.IsHidden,
                    column.OutlineLevel,
                    column.OriginalExcelWidth
                }),
                Cells = sheet.Cells.Select(cell => new
                {
                    cell.Row,
                    cell.Column,
                    cell.Address,
                    cell.DisplayText
                }),
                sheet.MergedRanges,
                sheet.Images,
                sheet.PageSetup,
                sheet.HeaderFooter,
                sheet.PrintArea,
                sheet.HorizontalPageBreaks,
                sheet.VerticalPageBreaks
            })
        };

    public static object CreatePdfPreparationPayload(ReportWorkbook workbook, object renderPlan) =>
        new
        {
            Workbook = CreateWorkbookPayload(workbook),
            RenderPlan = renderPlan,
            Environment = new
            {
                OperatingSystem = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
                Architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(),
                Culture = CultureInfo.CurrentCulture.Name,
                Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription
            }
        };
}
