namespace OysterReport;

using ClosedXML.Excel;

// 1 シートに対するプレースホルダ置換・行操作を提供する。
public sealed class TemplateSheet
{
#pragma warning disable IDE0032
    private readonly IXLWorksheet worksheet;
#pragma warning restore IDE0032

    internal TemplateSheet(IXLWorksheet ws)
    {
        worksheet = ws;
    }

    // シート名。
    public string Name => worksheet.Name;

    // 内部の ClosedXML ワークシート。
#pragma warning disable IDE0032
    internal IXLWorksheet UnderlyingWorksheet => worksheet;
#pragma warning restore IDE0032

    // マーカー名を指定してプレースホルダを置換する。
    public int ReplacePlaceholder(string markerName, string value)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        ArgumentNullException.ThrowIfNull(value);
        var placeholder = "{{" + markerName + "}}";
        var count = 0;

        foreach (var cell in worksheet.CellsUsed())
        {
            var text = cell.GetString();
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                cell.Value = text.Replace(placeholder, value, StringComparison.Ordinal);
                count++;
            }
        }

        return count;
    }

    // 辞書で一括置換する。
    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        ArgumentNullException.ThrowIfNull(values);
        var count = 0;
        foreach (var (key, value) in values)
        {
            count += ReplacePlaceholder(key, value ?? string.Empty);
        }
        return count;
    }

    // 行番号で単一行を取得する (1-based)。
    public SheetRow GetRow(int row) => new(worksheet, row);

    // 行番号で行範囲を取得する (1-based, inclusive)。
    public SheetRowRange GetRows(int startRow, int endRow) => new(worksheet, startRow, endRow);

    // マーカー名で行を検索して取得する。
    public SheetRow FindRow(string markerName)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        var placeholder = "{{" + markerName + "}}";
        foreach (var cell in worksheet.CellsUsed())
        {
            if (cell.GetString().Contains(placeholder, StringComparison.Ordinal))
            {
                return new SheetRow(worksheet, cell.Address.RowNumber);
            }
        }
        throw new InvalidOperationException($"Marker '{markerName}' not found in sheet '{Name}'.");
    }

    // マーカー名で行範囲を検索して取得する。プレースホルダを含む連続行範囲を自動検出する。
    public SheetRowRange FindRows(string markerName)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        var placeholder = "{{" + markerName + "}}";
        var markerRow = -1;

        foreach (var cell in worksheet.CellsUsed())
        {
            if (cell.GetString().Contains(placeholder, StringComparison.Ordinal))
            {
                markerRow = cell.Address.RowNumber;
                break;
            }
        }

        if (markerRow < 0)
        {
            throw new InvalidOperationException($"Marker '{markerName}' not found in sheet '{Name}'.");
        }

        var startRow = markerRow;
        var endRow = markerRow;

        while (RowContainsAnyPlaceholder(endRow + 1))
        {
            endRow++;
        }

        while (startRow > 1 && RowContainsAnyPlaceholder(startRow - 1))
        {
            startRow--;
        }

        return new SheetRowRange(worksheet, startRow, endRow);
    }

    // 指定範囲の行を削除する (1-based, inclusive)。
    public void DeleteRows(int startRow, int endRow)
    {
        worksheet.Rows(startRow, endRow).Delete();
    }

    private bool RowContainsAnyPlaceholder(int rowNum)
    {
        if (rowNum < 1)
        {
            return false;
        }

        var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        for (var col = 1; col <= lastCol; col++)
        {
            var text = worksheet.Cell(rowNum, col).GetString();
            if (text.Contains("{{", StringComparison.Ordinal) && text.Contains("}}", StringComparison.Ordinal))
            {
                return true;
            }
        }
        return false;
    }
}
