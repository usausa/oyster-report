namespace OysterReport;

using ClosedXML.Excel;

/// <summary>
/// 1 シートに対するプレースホルダ置換・行操作を提供する。
/// </summary>
public sealed class TemplateSheet
{
    private readonly IXLWorksheet _worksheet;

    internal TemplateSheet(IXLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    /// <summary>シート名。</summary>
    public string Name => _worksheet.Name;

    /// <summary>内部の ClosedXML ワークシート（上級者向け）。</summary>
    [CLSCompliant(false)]
    public IXLWorksheet UnderlyingWorksheet => _worksheet;

    /// <summary>マーカー名を指定してプレースホルダを置換する。</summary>
    public int ReplacePlaceholder(string markerName, string value)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        ArgumentNullException.ThrowIfNull(value);
        var placeholder = "{{" + markerName + "}}";
        var count = 0;

        foreach (var cell in _worksheet.CellsUsed())
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

    /// <summary>辞書で一括置換する。</summary>
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

    /// <summary>行番号で単一行を取得する (1-based)。</summary>
    public SheetRow GetRow(int row) => new(_worksheet, row);

    /// <summary>行番号で行範囲を取得する (1-based, inclusive)。</summary>
    public SheetRowRange GetRows(int startRow, int endRow) => new(_worksheet, startRow, endRow);

    /// <summary>マーカー名で行を検索して取得する。</summary>
    public SheetRow FindRow(string markerName)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        var placeholder = "{{" + markerName + "}}";
        foreach (var cell in _worksheet.CellsUsed())
        {
            if (cell.GetString().Contains(placeholder, StringComparison.Ordinal))
            {
                return new SheetRow(_worksheet, cell.Address.RowNumber);
            }
        }
        throw new InvalidOperationException($"Marker '{markerName}' not found in sheet '{Name}'.");
    }

    /// <summary>マーカー名で行範囲を検索して取得する。プレースホルダを含む連続行範囲を自動検出する。</summary>
    public SheetRowRange FindRows(string markerName)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        var placeholder = "{{" + markerName + "}}";
        var markerRow = -1;

        foreach (var cell in _worksheet.CellsUsed())
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

        return new SheetRowRange(_worksheet, startRow, endRow);
    }

    /// <summary>指定範囲の行を削除する (1-based, inclusive)。</summary>
    public void DeleteRows(int startRow, int endRow)
    {
        _worksheet.Rows(startRow, endRow).Delete();
    }

    private bool RowContainsAnyPlaceholder(int rowNum)
    {
        if (rowNum < 1) return false;
        var lastCol = _worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        for (var col = 1; col <= lastCol; col++)
        {
            var text = _worksheet.Cell(rowNum, col).GetString();
            if (text.Contains("{{", StringComparison.Ordinal) && text.Contains("}}", StringComparison.Ordinal))
            {
                return true;
            }
        }
        return false;
    }
}
