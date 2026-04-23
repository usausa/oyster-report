// Reads table parts attached to a worksheet (tables/tableN.xml) and emits (range, theme, showHeader, showTotals, showStripes).
// Caller applies row-stripe fills via TableStyleCatalog (in OysterReport.Internal).

namespace OysterReport.Prototype;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using OysterReport.Internal;

internal static class TableLoader
{
    public static IEnumerable<TableInfo> Load(WorksheetPart worksheetPart)
    {
        foreach (var part in worksheetPart.TableDefinitionParts)
        {
            var table = part.Table;
            if (table is null || table.Reference?.Value is not { } refStr)
            {
                continue;
            }

            var range = ParseRangeRef(refStr);
            var themeName = table.TableStyleInfo?.Name?.Value;
            var showRowStripes = table.TableStyleInfo?.ShowRowStripes?.Value ?? false;
            var showHeader = table.HeaderRowCount is null || table.HeaderRowCount.Value > 0;
            var showTotals = table.TotalsRowCount is not null && table.TotalsRowCount.Value > 0;

            yield return new TableInfo(range, themeName, showRowStripes, showHeader, showTotals);
        }
    }

    private static ReportRange ParseRangeRef(string reference)
    {
        var colonIdx = reference.IndexOf(':', StringComparison.Ordinal);
        if (colonIdx < 0)
        {
            var (row, col) = AddressHelper.ParseAddress(reference);
            return new ReportRange { StartRow = row, StartColumn = col, EndRow = row, EndColumn = col };
        }

        var (r1, c1) = AddressHelper.ParseAddress(reference[..colonIdx]);
        var (r2, c2) = AddressHelper.ParseAddress(reference[(colonIdx + 1)..]);
        return new ReportRange
        {
            StartRow = Math.Min(r1, r2),
            StartColumn = Math.Min(c1, c2),
            EndRow = Math.Max(r1, r2),
            EndColumn = Math.Max(c1, c2)
        };
    }

    internal sealed record TableInfo(ReportRange Range, string? ThemeName, bool ShowRowStripes, bool ShowHeader, bool ShowTotals);
}
