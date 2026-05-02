namespace OysterReport.Internal.OpenXml;

using DocumentFormat.OpenXml.Packaging;

internal sealed record TableInfo(
    ReportRange Range,
    string? ThemeName,
    bool ShowRowStripes,
    bool ShowHeader,
    bool ShowTotals);

internal static class TableLoader
{
    public static IEnumerable<TableInfo> Load(WorksheetPart worksheetPart)
    {
        foreach (var part in worksheetPart.TableDefinitionParts)
        {
            var table = part.Table;
            if (table?.Reference?.Value is not { } refStr)
            {
                continue;
            }

            var range = ParseRangeRef(refStr);
            var themeName = table.TableStyleInfo?.Name?.Value;
            var showRowStripes = table.TableStyleInfo?.ShowRowStripes?.Value ?? false;
            var showHeader = (table.HeaderRowCount is null) || (table.HeaderRowCount.Value > 0);
            var showTotals = (table.TotalsRowCount is not null) && (table.TotalsRowCount.Value > 0);

            yield return new TableInfo(range, themeName, showRowStripes, showHeader, showTotals);
        }
    }

    private static ReportRange ParseRangeRef(string reference)
    {
        var index = reference.IndexOf(':', StringComparison.Ordinal);
        if (index < 0)
        {
            AddressHelper.ParseAddress(reference, out var row, out var col);
            return new ReportRange { StartRow = row, StartColumn = col, EndRow = row, EndColumn = col };
        }

        AddressHelper.ParseAddress(reference[..index], out var r1, out var c1);
        AddressHelper.ParseAddress(reference[(index + 1)..], out var r2, out var c2);
        return new ReportRange
        {
            StartRow = Math.Min(r1, r2),
            StartColumn = Math.Min(c1, c2),
            EndRow = Math.Max(r1, r2),
            EndColumn = Math.Max(c1, c2)
        };
    }
}
