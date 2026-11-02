namespace OysterReport;

using OysterReport.Internal;

public sealed class TemplateSheet
{
    private readonly ReportWorkbook workbook;

    internal ReportSheet UnderlyingSheet { get; }

    internal ReportMeasurementProfile WorkbookMeasurementProfile => workbook.MeasurementProfile;

    internal ReportMetadata WorkbookMetadata => workbook.Metadata;

    public string Name => UnderlyingSheet.Name;

    //--------------------------------------------------------------------------------
    // Constructor
    //--------------------------------------------------------------------------------

    internal TemplateSheet(ReportWorkbook workbook, ReportSheet sheet)
    {
        this.workbook = workbook;
        UnderlyingSheet = sheet;
    }

    //--------------------------------------------------------------------------------
    // Row
    //--------------------------------------------------------------------------------

    public TemplateRow GetRow(int row) => new(UnderlyingSheet, row);

    public TemplateRowRange GetRows(int startRow, int endRow) => new(UnderlyingSheet, startRow, endRow);

    public TemplateRow FindRow(string marker)
    {
        var placeholder = "{{" + marker + "}}";
        foreach (var cell in UnderlyingSheet.Cells)
        {
            if (cell.DisplayText.Contains(placeholder, StringComparison.Ordinal))
            {
                return new TemplateRow(UnderlyingSheet, cell.Row);
            }
        }

        throw new InvalidOperationException($"Marker not found in sheet. marker=[{marker}], sheet=[{Name}]");
    }

    public TemplateRowRange FindRows(string marker)
    {
        var placeholder = "{{" + marker + "}}";
        var markerRow = -1;
        foreach (var cell in UnderlyingSheet.Cells)
        {
            if (cell.DisplayText.Contains(placeholder, StringComparison.Ordinal))
            {
                markerRow = cell.Row;
                break;
            }
        }

        if (markerRow < 0)
        {
            throw new InvalidOperationException($"Marker not found in sheet. marker=[{marker}], sheet=[{Name}]");
        }

        var startRow = markerRow;
        var endRow = markerRow;

        while (RowContainsAnyPlaceholder(endRow + 1))
        {
            endRow++;
        }

        while ((startRow > 1) && RowContainsAnyPlaceholder(startRow - 1))
        {
            startRow--;
        }

        return new TemplateRowRange(UnderlyingSheet, startRow, endRow);
    }

    private bool RowContainsAnyPlaceholder(int rowNum)
    {
        if (rowNum < 1)
        {
            return false;
        }

        foreach (var cell in UnderlyingSheet.CellsInStorageOrder)
        {
            if (cell.Row != rowNum)
            {
                continue;
            }

            var text = cell.DisplayText;
            if (text.Contains("{{", StringComparison.Ordinal) && text.Contains("}}", StringComparison.Ordinal))
            {
                return true;
            }
        }
        return false;
    }

    public void DeleteRows(int startRow, int endRow)
    {
        UnderlyingSheet.DeleteRows(startRow, endRow);
    }

    //--------------------------------------------------------------------------------
    // Edit
    //--------------------------------------------------------------------------------

    public int ReplacePlaceholder(string marker, string value)
    {
        var placeholder = "{{" + marker + "}}";
        var count = 0;

        foreach (var cell in UnderlyingSheet.CellsInStorageOrder)
        {
            var text = cell.DisplayText;
            if (text.Contains(placeholder, StringComparison.Ordinal))
            {
                var replaced = text.Replace(placeholder, value, StringComparison.Ordinal);
                SetCellText(cell, replaced);
                count++;
            }
        }

        return count + PlaceholderReplacer.ReplaceInHeaderFooter(UnderlyingSheet, marker, value);
    }

    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        return PlaceholderReplacer.ReplaceInRows(UnderlyingSheet, 1, Int32.MaxValue, values) +
               PlaceholderReplacer.ReplaceInHeaderFooter(UnderlyingSheet, values);
    }

    public int ReplacePlaceholders(IEnumerable<IReadOnlyDictionary<string, string?>> rows)
    {
        var positions = new Dictionary<string, (int Row, int Column)>();
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

    public (int Row, int Column) FindMarkerPosition(string marker)
    {
        if (!TryFindMarkerPosition(marker, out var position))
        {
            throw new InvalidOperationException($"Marker not found in sheet. marker=[{marker}], sheet=[{Name}]");
        }
        return position;
    }

    public bool TryFindMarkerPosition(string marker, out (int Row, int Column) position)
    {
        var placeholder = "{{" + marker + "}}";
        foreach (var cell in UnderlyingSheet.Cells)
        {
            if (cell.DisplayText.Contains(placeholder, StringComparison.Ordinal))
            {
                position = (cell.Row, cell.Column);
                return true;
            }
        }

        position = default;
        return false;
    }

    public string GetCellText(int row, int column)
    {
        var cell = UnderlyingSheet.FindCell(row, column);
        return cell is null ? string.Empty : cell.DisplayText;
    }

    public void SetCellValue(int row, int column, string? value)
    {
        SetCellValueCore(row, column, value ?? string.Empty);
    }

    private void SetCellValueCore(int row, int column, string value)
    {
        var cell = UnderlyingSheet.FindCell(row, column);
        if (cell is null)
        {
            cell = new ReportCell { Row = row, Column = column };
            UnderlyingSheet.AddCell(cell);
            UnderlyingSheet.EnsureRowDefinition(row);
        }
        SetCellText(cell, value);
    }

    internal static void SetCellText(ReportCell cell, string value)
    {
        cell.Value = new ReportCellValue
        {
            Kind = CellValueKind.Text,
            RawValue = value
        };
        cell.DisplayText = value;
    }

    //--------------------------------------------------------------------------------
    // Image
    //--------------------------------------------------------------------------------

    public int ReplacePlaceholderWithImage(string marker, ReadOnlyMemory<byte> imageBytes, ReportImageFit fit = ReportImageFit.Contain)
    {
        return PlaceholderReplacer.ReplaceWithImageInRows(UnderlyingSheet, 1, Int32.MaxValue, marker, imageBytes, fit);
    }

    public void SetCellImage(int row, int column, ReadOnlyMemory<byte> imageBytes, ReportImageFit fit = ReportImageFit.Contain)
    {
        var merge = UnderlyingSheet.FindCell(row, column)?.Merge;
        AddCellImage(UnderlyingSheet, row, column, merge, AddressHelper.ToAddress(row, column), imageBytes, fit);
    }

    public int ReplaceImage(string name, ReadOnlyMemory<byte> imageBytes, ReportImageFit fit = ReportImageFit.Contain)
    {
        var count = 0;
        foreach (var image in UnderlyingSheet.Images)
        {
            if (String.Equals(image.Name, name, StringComparison.Ordinal))
            {
                image.ImageBytes = imageBytes;
                image.Fit = fit;
                count++;
            }
        }

        return count;
    }

    public int RemoveImage(string name)
    {
        return UnderlyingSheet.RemoveImages(x => String.Equals(x.Name, name, StringComparison.Ordinal));
    }

    internal static void AddCellImage(ReportSheet sheet, int row, int column, ReportMergeInfo? merge, string name, ReadOnlyMemory<byte> imageBytes, ReportImageFit fit)
    {
        var range = merge?.Range ?? new ReportRange { StartRow = row, StartColumn = column, EndRow = row, EndColumn = column };
        sheet.AddImage(new ReportImage
        {
            Name = name,
            FromCellAddress = AddressHelper.ToAddress(range.StartRow, range.StartColumn),
            ToCellAddress = AddressHelper.ToAddress(range.EndRow, range.EndColumn),
            InCell = true,
            Fit = fit,
            ImageBytes = imageBytes
        });
    }
}
