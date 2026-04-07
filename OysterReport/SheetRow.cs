namespace OysterReport;

using ClosedXML.Excel;

/// <summary>
/// シート上の 1 行を表す軽量ハンドル。コピー挿入・プレースホルダ置換・削除を提供する。
/// </summary>
public sealed class SheetRow
{
    private readonly IXLWorksheet _worksheet;

    internal SheetRow(IXLWorksheet worksheet, int rowNumber)
    {
        _worksheet = worksheet;
        RowNumber = rowNumber;
    }

    /// <summary>この行の行番号 (1-based)。</summary>
    public int RowNumber { get; }

    /// <summary>
    /// この行のコピーを直下に挿入し、挿入された新しい行を返す。
    /// フロー B（行番号を進めながら処理する方式）で使用する。
    /// </summary>
    public SheetRow InsertCopyBelow()
    {
        return InsertCopyAfter(this);
    }

    /// <summary>
    /// この行の内容をコピーし、afterRow の直下に挿入する。挿入された新しい行を返す。
    /// コピー元は this、挿入位置は afterRow の直下。フロー A（テンプレートのコピーを追加していく方式）で使用する。
    /// </summary>
    public SheetRow InsertCopyAfter(SheetRow afterRow)
    {
        ArgumentNullException.ThrowIfNull(afterRow);
        var newRowNumber = afterRow.RowNumber + 1;

        _worksheet.Row(newRowNumber).InsertRowsAbove(1);
        var insertedRow = _worksheet.Row(newRowNumber);

        // this (コピー元) の行番号を再計算: 挿入位置が this より上なら +1 シフト
        var sourceRowNum = (newRowNumber <= RowNumber) ? RowNumber + 1 : RowNumber;

        var lastColumn = _worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        for (var col = 1; col <= lastColumn; col++)
        {
            var srcCell = _worksheet.Cell(sourceRowNum, col);
            var dstCell = _worksheet.Cell(newRowNumber, col);
            dstCell.Value = srcCell.Value;
            dstCell.Style = srcCell.Style;
        }

        insertedRow.Height = _worksheet.Row(sourceRowNum).Height;

        return new SheetRow(_worksheet, newRowNumber);
    }

    /// <summary>この行内のプレースホルダを置換する。</summary>
    public int ReplacePlaceholder(string markerName, string value)
    {
        ArgumentNullException.ThrowIfNull(markerName);
        ArgumentNullException.ThrowIfNull(value);
        var placeholder = "{{" + markerName + "}}";
        var count = 0;
        var lastColumn = _worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        for (var col = 1; col <= lastColumn; col++)
        {
            var cell = _worksheet.Cell(RowNumber, col);
            var text = cell.GetString();
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                cell.Value = text.Replace(placeholder, value, StringComparison.Ordinal);
                count++;
            }
        }

        return count;
    }

    /// <summary>この行内のプレースホルダを辞書で一括置換する。</summary>
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

    /// <summary>この行を削除する。後続行は自動的に上にシフトされる。</summary>
    public void Delete()
    {
        _worksheet.Row(RowNumber).Delete();
    }
}
