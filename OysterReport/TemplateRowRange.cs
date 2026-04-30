namespace OysterReport;

using OysterReport.Internal;

public sealed class TemplateRowRange
{
    private readonly ReportSheet sheet;

    public int StartRow { get; private set; }

    public int EndRow { get; private set; }

    public int RowCount => EndRow - StartRow + 1;

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    internal TemplateRowRange(ReportSheet sheet, int startRow, int endRow)
    {
        this.sheet = sheet;
        StartRow = startRow;
        EndRow = endRow;
    }

    //--------------------------------------------------------------------------------
    // Edit
    //--------------------------------------------------------------------------------

    public TemplateRowRange InsertCopyBelow()
    {
        return InsertCopyAfter(this);
    }

    public TemplateRowRange InsertCopyAfter(TemplateRowRange afterRange)
    {
        var newStartRow = afterRange.EndRow + 1;
        var rowCount = RowCount;

        var sourceStartRow = (newStartRow <= StartRow) ? StartRow + rowCount : StartRow;

        sheet.InsertEmptyRowsAt(newStartRow, rowCount);

        if (newStartRow <= StartRow)
        {
            StartRow += rowCount;
            EndRow += rowCount;
        }

        for (var offset = 0; offset < rowCount; offset++)
        {
            sheet.CopyRowContent(sourceStartRow + offset, newStartRow + offset);
        }

        return new TemplateRowRange(sheet, newStartRow, newStartRow + rowCount - 1);
    }

    public void Delete()
    {
        sheet.DeleteRows(StartRow, EndRow);
    }

    public int ReplacePlaceholder(string markerName, string value)
    {
        var placeholder = "{{" + markerName + "}}";
        var count = 0;

        foreach (var cell in sheet.Cells)
        {
            if (cell.Row < StartRow || cell.Row > EndRow)
            {
                continue;
            }

            var text = cell.DisplayText;
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                var replaced = text.Replace(placeholder, value, StringComparison.Ordinal);
                TemplateSheet.SetCellText(cell, replaced);
                count++;
            }
        }

        return count;
    }

    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        var count = 0;
        foreach (var (key, value) in values)
        {
            count += ReplacePlaceholder(key, value ?? string.Empty);
        }
        return count;
    }

    public int ReplacePlaceholders(IEnumerable<IReadOnlyDictionary<string, string?>> rows)
    {
        var positions = new Dictionary<string, (int Row, int Column)>(StringComparer.Ordinal);
        var rowIndex = 0;
        var count = 0;
        foreach (var entry in rows)
        {
            foreach (var (key, value) in entry)
            {
                if (!positions.TryGetValue(key, out var pos))
                {
                    if (!TryFindMarkerPosition(key, out pos))
                    {
                        continue;
                    }
                    positions[key] = pos;
                }

                SetCellValueCore(pos.Row + rowIndex, pos.Column, value ?? string.Empty);
                count++;
            }

            rowIndex++;
        }

        return count;
    }

    private bool TryFindMarkerPosition(string marker, out (int Row, int Column) position)
    {
        var placeholder = "{{" + marker + "}}";
        foreach (var cell in sheet.Cells)
        {
            if (cell.Row < StartRow || cell.Row > EndRow)
            {
                continue;
            }

            if (cell.DisplayText.Contains(placeholder, StringComparison.Ordinal))
            {
                position = (cell.Row, cell.Column);
                return true;
            }
        }

        position = default;
        return false;
    }

    private void SetCellValueCore(int row, int column, string value)
    {
        var cell = sheet.FindCell(row, column);
        if (cell is null)
        {
            cell = new ReportCell { Row = row, Column = column };
            sheet.AddCell(cell);
        }
        TemplateSheet.SetCellText(cell, value);
    }
}
