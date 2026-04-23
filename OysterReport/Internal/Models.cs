namespace OysterReport.Internal;

using System.Diagnostics.CodeAnalysis;

// ReportWorkbook
// ├── ReportMetadata                 Template name
// ├── ReportMeasurementProfile       Font metrics for column width calculation
// └── ReportSheet[]
//     ├── ReportRow[]                Row height, visibility, and outline level
//     ├── ReportColumn[]             Column width and visibility
//     ├── ReportCell[]
//     │   ├── ReportCellValue       Typed value
//     │   ├── ReportCellStyle
//     │   │   ├── ReportFont
//     │   │   ├── ReportFill
//     │   │   ├── ReportBorders
//     │   │   └── ReportAlignment
//     │   └── ReportMergeInfo?      Merge owner information
//     ├── ReportMergedRange[]        Merged cell ranges
//     ├── ReportImage[]              Embedded images
//     ├── ReportPageSetup            Paper, margins, and centering
//     ├── ReportHeaderFooter         Header/footer text
//     ├── ReportPrintArea?           Print area
//     └── ReportPageBreak[]          Horizontal/vertical page breaks

//--------------------------------------------------------------------------------
// Metadata
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty;
}

//--------------------------------------------------------------------------------
// Measurement profile
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportMeasurementProfile
{
    public double MaxDigitWidth { get; init; } = 7d;

    public double ColumnWidthAdjustment { get; init; } = 1d;
}

//--------------------------------------------------------------------------------
// Cell value and style
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportCellValue
{
    public CellValueKind Kind { get; init; } = CellValueKind.Blank;

    // [MEMO]: Use the typed source value when adding value-aware formatting or placeholder features.
    public object? RawValue { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportFont
{
    public string Name { get; init; } = "Arial";

    public double Size { get; init; } = 11d;

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public bool Underline { get; init; }

    public bool Strikeout { get; init; }

    public string ColorHex { get; init; } = "#FF000000";
}

[ExcludeFromCodeCoverage]
internal sealed record ReportFill
{
    public string BackgroundColorHex { get; init; } = "#00000000";
}

[ExcludeFromCodeCoverage]
internal sealed record ReportBorder
{
    public BorderLineStyle Style { get; init; } = BorderLineStyle.None;

    public string ColorHex { get; init; } = "#FF000000";

    public double Width { get; init; } = 0.5d;
}

[ExcludeFromCodeCoverage]
internal sealed record ReportBorders
{
    public ReportBorder Left { get; init; } = new();

    public ReportBorder Top { get; init; } = new();

    public ReportBorder Right { get; init; } = new();

    public ReportBorder Bottom { get; init; } = new();
}

[ExcludeFromCodeCoverage]
internal sealed record ReportAlignment
{
    public HorizontalAlignment Horizontal { get; init; } = HorizontalAlignment.General;

    public VerticalAlignment Vertical { get; init; } = VerticalAlignment.Top;
}

[ExcludeFromCodeCoverage]
internal sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new();

    public ReportFill Fill { get; init; } = new();

    public ReportBorders Borders { get; init; } = new();

    public ReportAlignment Alignment { get; init; } = new();

    public bool WrapText { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportMergeInfo
{
    // [MEMO] Retained for debugging and for future features that need to identify the owner from a non-owner merged cell.
    public string OwnerCellAddress { get; init; } = string.Empty;

    public ReportRange Range { get; init; }
}

//--------------------------------------------------------------------------------
// Page setup
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed record ReportPageBreak
{
    // [MEMO] Not used by the current single-page rendering flow, but kept for future multipage support together with HorizontalPageBreaks / VerticalPageBreaks.
    public int Index { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportPrintArea
{
    public ReportRange Range { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportHeaderFooter
{
    // [MEMO]: Apply Excel's header/footer margin alignment rules during PDF rendering.
    public bool AlignWithMargins { get; init; } = true;

    public bool DifferentFirst { get; init; }

    public bool DifferentOddEven { get; init; }

    // [MEMO]: Apply Excel's header/footer scaling rule during PDF rendering.
    public bool ScaleWithDocument { get; init; } = true;

    public string? OddHeader { get; init; }

    public string? OddFooter { get; init; }

    public string? EvenHeader { get; init; }

    public string? EvenFooter { get; init; }

    public string? FirstHeader { get; init; }

    public string? FirstFooter { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed record ReportPageSetup
{
    public PaperSize PaperSize { get; init; } = PaperSize.A4Paper;

    public PageOrientation Orientation { get; init; } = PageOrientation.Default;

    public ReportThickness Margins { get; init; } = new() { Left = 36d, Top = 36d, Right = 36d, Bottom = 36d };

    public double HeaderMarginPoint { get; init; } = 18d;

    public double FooterMarginPoint { get; init; } = 18d;

    // [MEMO]: Apply Excel print scaling when multipage fit/scaling support is implemented.
    public int ScalePercent { get; init; } = 100;

    // [MEMO]: Apply Excel fit-to-page width scaling when multipage fit/scaling support is implemented.
    public int? FitToPagesWide { get; init; }

    // [MEMO]: Apply Excel fit-to-page height scaling when multipage fit/scaling support is implemented.
    public int? FitToPagesTall { get; init; }

    public bool CenterHorizontally { get; init; }

    public bool CenterVertically { get; init; }
}

//--------------------------------------------------------------------------------
// Sheet structure
//--------------------------------------------------------------------------------

[ExcludeFromCodeCoverage]
internal sealed class ReportRow
{
    public int Index { get; set; }

    public double HeightPoint { get; set; }

    public double TopPoint { get; set; }

    public bool IsHidden { get; init; }

    public int OutlineLevel { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportColumn
{
    public int Index { get; set; }

    public double WidthPoint { get; init; }

    public double LeftPoint { get; set; }

    public bool IsHidden { get; init; }

    public int OutlineLevel { get; init; }

    public double OriginalExcelWidth { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportMergedRange
{
    public ReportRange Range { get; set; }

    public string OwnerCellAddress => AddressHelper.ToAddress(Range.StartRow, Range.StartColumn);
}

[ExcludeFromCodeCoverage]
internal sealed class ReportImage
{
    public string Name { get; init; } = string.Empty;

    public string FromCellAddress { get; set; } = string.Empty;

    public string? ToCellAddress { get; set; }

    public ReportOffset Offset { get; init; }

    public double WidthPoint { get; init; }

    public double HeightPoint { get; init; }

    public ReadOnlyMemory<byte> ImageBytes { get; init; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportCell
{
    public int Row { get; set; }

    public int Column { get; set; }

    public string Address => AddressHelper.ToAddress(Row, Column);

    public ReportCellValue Value { get; set; } = new();

    public string DisplayText { get; set; } = string.Empty;

    public ReportCellStyle Style { get; set; } = new();

    public ReportMergeInfo? Merge { get; set; }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportSheet
{
    private readonly List<ReportRow> rows = [];
    private readonly List<ReportColumn> columns = [];
    private readonly List<ReportCell> cells = [];
    private readonly List<ReportMergedRange> mergedRanges = [];
    private readonly List<ReportImage> images = [];
    private readonly List<ReportPageBreak> horizontalPageBreaks = [];
    private readonly List<ReportPageBreak> verticalPageBreaks = [];

    public string Name { get; set; } = string.Empty;

    public ReportRange UsedRange { get; set; } = new() { StartRow = 1, StartColumn = 1, EndRow = 1, EndColumn = 1 };

    public IReadOnlyList<ReportRow> Rows => rows;

    public IReadOnlyList<ReportColumn> Columns => columns;

    public IReadOnlyList<ReportCell> Cells => cells;

    public IReadOnlyList<ReportMergedRange> MergedRanges => mergedRanges;

    public IReadOnlyList<ReportImage> Images => images;

    public ReportPageSetup PageSetup { get; set; } = new();

    public ReportHeaderFooter HeaderFooter { get; set; } = new();

    public ReportPrintArea? PrintArea { get; set; }

    public IReadOnlyList<ReportPageBreak> HorizontalPageBreaks => horizontalPageBreaks;

    public IReadOnlyList<ReportPageBreak> VerticalPageBreaks => verticalPageBreaks;

    public bool ShowGridLines { get; set; }

    public void AddRowDefinition(ReportRow row) => rows.Add(row);

    public void AddColumnDefinition(ReportColumn column) => columns.Add(column);

    public void AddCell(ReportCell cell) => cells.Add(cell);

    public void AddMergedRange(ReportMergedRange range) => mergedRanges.Add(range);

    public void AddImage(ReportImage image) => images.Add(image);

    public void AddHorizontalPageBreak(ReportPageBreak pageBreak) => horizontalPageBreaks.Add(pageBreak);

    public void AddVerticalPageBreak(ReportPageBreak pageBreak) => verticalPageBreaks.Add(pageBreak);

    public void RecalculateLayout()
    {
        rows.Sort(static (a, b) => a.Index.CompareTo(b.Index));
        columns.Sort(static (a, b) => a.Index.CompareTo(b.Index));
        cells.Sort(static (a, b) =>
        {
            var r = a.Row.CompareTo(b.Row);
            return r != 0 ? r : a.Column.CompareTo(b.Column);
        });

        var top = 0d;
        foreach (var row in rows)
        {
            row.TopPoint = top;
            top += row.HeightPoint;
        }

        var left = 0d;
        foreach (var column in columns)
        {
            column.LeftPoint = left;
            left += column.WidthPoint;
        }
    }

    //--------------------------------------------------------------------------------
    // Mutation APIs for templates
    //--------------------------------------------------------------------------------

    public ReportRow? GetRowDefinition(int row)
    {
        foreach (var r in rows)
        {
            if (r.Index == row)
            {
                return r;
            }
        }
        return null;
    }

    public ReportCell? FindCell(int row, int column)
    {
        foreach (var cell in cells)
        {
            if (cell.Row == row && cell.Column == column)
            {
                return cell;
            }
        }
        return null;
    }

    // Inserts `count` blank rows at `insertAtRow`, shifting existing rows/cells/merges/images/breaks down.
    public void InsertEmptyRowsAt(int insertAtRow, int count)
    {
        if (count <= 0)
        {
            return;
        }

        foreach (var row in rows)
        {
            if (row.Index >= insertAtRow)
            {
                row.Index += count;
            }
        }

        for (var i = 0; i < count; i++)
        {
            rows.Add(new ReportRow
            {
                Index = insertAtRow + i,
                HeightPoint = 15d
            });
        }

        foreach (var cell in cells)
        {
            if (cell.Row >= insertAtRow)
            {
                cell.Row += count;
                if (cell.Merge is not null)
                {
                    var mr = cell.Merge.Range;
                    cell.Merge = new ReportMergeInfo
                    {
                        OwnerCellAddress = AddressHelper.ToAddress(mr.StartRow + count, mr.StartColumn),
                        Range = new ReportRange
                        {
                            StartRow = mr.StartRow + count,
                            StartColumn = mr.StartColumn,
                            EndRow = mr.EndRow + count,
                            EndColumn = mr.EndColumn
                        }
                    };
                }
            }
            else if (cell.Merge is not null && cell.Merge.Range.EndRow >= insertAtRow)
            {
                var mr = cell.Merge.Range;
                cell.Merge = new ReportMergeInfo
                {
                    OwnerCellAddress = cell.Merge.OwnerCellAddress,
                    Range = new ReportRange
                    {
                        StartRow = mr.StartRow,
                        StartColumn = mr.StartColumn,
                        EndRow = mr.EndRow + count,
                        EndColumn = mr.EndColumn
                    }
                };
            }
        }

        foreach (var mr in mergedRanges)
        {
            var r = mr.Range;
            if (r.StartRow >= insertAtRow)
            {
                mr.Range = new ReportRange
                {
                    StartRow = r.StartRow + count,
                    StartColumn = r.StartColumn,
                    EndRow = r.EndRow + count,
                    EndColumn = r.EndColumn
                };
            }
            else if (r.EndRow >= insertAtRow)
            {
                mr.Range = new ReportRange
                {
                    StartRow = r.StartRow,
                    StartColumn = r.StartColumn,
                    EndRow = r.EndRow + count,
                    EndColumn = r.EndColumn
                };
            }
        }

        ShiftImageRows(insertAtRow, count);

        for (var i = 0; i < horizontalPageBreaks.Count; i++)
        {
            if (horizontalPageBreaks[i].Index >= insertAtRow)
            {
                horizontalPageBreaks[i] = new ReportPageBreak { Index = horizontalPageBreaks[i].Index + count };
            }
        }

        if (PrintArea is not null)
        {
            var pa = PrintArea.Range;
            if (pa.StartRow >= insertAtRow)
            {
                PrintArea = new ReportPrintArea
                {
                    Range = new ReportRange
                    {
                        StartRow = pa.StartRow + count,
                        StartColumn = pa.StartColumn,
                        EndRow = pa.EndRow + count,
                        EndColumn = pa.EndColumn
                    }
                };
            }
            else if (pa.EndRow >= insertAtRow)
            {
                PrintArea = new ReportPrintArea
                {
                    Range = new ReportRange
                    {
                        StartRow = pa.StartRow,
                        StartColumn = pa.StartColumn,
                        EndRow = pa.EndRow + count,
                        EndColumn = pa.EndColumn
                    }
                };
            }
        }

        if (UsedRange.StartRow >= insertAtRow || UsedRange.EndRow >= insertAtRow)
        {
            var u = UsedRange;
            UsedRange = new ReportRange
            {
                StartRow = u.StartRow >= insertAtRow ? u.StartRow + count : u.StartRow,
                StartColumn = u.StartColumn,
                EndRow = u.EndRow >= insertAtRow ? u.EndRow + count : u.EndRow,
                EndColumn = u.EndColumn
            };
        }

        RecalculateLayout();
    }

    // Deletes rows in [startRow, endRow] (inclusive) and shifts everything below up.
    public void DeleteRows(int startRow, int endRow)
    {
        if (endRow < startRow)
        {
            return;
        }

        var count = endRow - startRow + 1;

        rows.RemoveAll(r => r.Index >= startRow && r.Index <= endRow);
        foreach (var row in rows)
        {
            if (row.Index > endRow)
            {
                row.Index -= count;
            }
        }

        cells.RemoveAll(c => c.Row >= startRow && c.Row <= endRow);
        foreach (var cell in cells)
        {
            if (cell.Row > endRow)
            {
                cell.Row -= count;
                if (cell.Merge is not null)
                {
                    var mr = cell.Merge.Range;
                    cell.Merge = new ReportMergeInfo
                    {
                        OwnerCellAddress = AddressHelper.ToAddress(Math.Max(1, mr.StartRow - count), mr.StartColumn),
                        Range = new ReportRange
                        {
                            StartRow = Math.Max(1, mr.StartRow - count),
                            StartColumn = mr.StartColumn,
                            EndRow = mr.EndRow - count,
                            EndColumn = mr.EndColumn
                        }
                    };
                }
            }
        }

        for (var i = mergedRanges.Count - 1; i >= 0; i--)
        {
            var mr = mergedRanges[i].Range;

            if (mr.EndRow < startRow)
            {
                continue;
            }

            if (mr.StartRow > endRow)
            {
                mergedRanges[i].Range = new ReportRange
                {
                    StartRow = mr.StartRow - count,
                    StartColumn = mr.StartColumn,
                    EndRow = mr.EndRow - count,
                    EndColumn = mr.EndColumn
                };
            }
            else if (mr.StartRow >= startRow && mr.EndRow <= endRow)
            {
                mergedRanges.RemoveAt(i);
            }
            else
            {
                var overlap = Math.Min(mr.EndRow, endRow) - Math.Max(mr.StartRow, startRow) + 1;
                var newStartRow = mr.StartRow < startRow ? mr.StartRow : startRow;
                var newEndRow = mr.EndRow - overlap;
                if (newEndRow < newStartRow)
                {
                    mergedRanges.RemoveAt(i);
                }
                else
                {
                    mergedRanges[i].Range = new ReportRange
                    {
                        StartRow = newStartRow,
                        StartColumn = mr.StartColumn,
                        EndRow = newEndRow,
                        EndColumn = mr.EndColumn
                    };
                }
            }
        }

        for (var i = images.Count - 1; i >= 0; i--)
        {
            var img = images[i];
            var fromRow = TryParseRow(img.FromCellAddress);
            if (fromRow is null)
            {
                continue;
            }

            if (fromRow.Value >= startRow && fromRow.Value <= endRow)
            {
                images.RemoveAt(i);
                continue;
            }

            if (fromRow.Value > endRow)
            {
                img.FromCellAddress = ReplaceRow(img.FromCellAddress, fromRow.Value - count);
                if (img.ToCellAddress is not null)
                {
                    var toRow = TryParseRow(img.ToCellAddress);
                    if (toRow is not null && toRow.Value > endRow)
                    {
                        img.ToCellAddress = ReplaceRow(img.ToCellAddress, toRow.Value - count);
                    }
                }
            }
        }

        for (var i = horizontalPageBreaks.Count - 1; i >= 0; i--)
        {
            var idx = horizontalPageBreaks[i].Index;
            if (idx >= startRow && idx <= endRow)
            {
                horizontalPageBreaks.RemoveAt(i);
            }
            else if (idx > endRow)
            {
                horizontalPageBreaks[i] = new ReportPageBreak { Index = idx - count };
            }
        }

        if (PrintArea is not null)
        {
            var pa = PrintArea.Range;
            if (pa.StartRow > endRow)
            {
                PrintArea = new ReportPrintArea
                {
                    Range = new ReportRange
                    {
                        StartRow = pa.StartRow - count,
                        StartColumn = pa.StartColumn,
                        EndRow = pa.EndRow - count,
                        EndColumn = pa.EndColumn
                    }
                };
            }
            else if (pa.EndRow >= startRow)
            {
                var overlap = Math.Min(pa.EndRow, endRow) - Math.Max(pa.StartRow, startRow) + 1;
                PrintArea = new ReportPrintArea
                {
                    Range = new ReportRange
                    {
                        StartRow = pa.StartRow,
                        StartColumn = pa.StartColumn,
                        EndRow = Math.Max(pa.StartRow, pa.EndRow - overlap),
                        EndColumn = pa.EndColumn
                    }
                };
            }
        }

        var u = UsedRange;
        if (u.StartRow > endRow)
        {
            UsedRange = new ReportRange
            {
                StartRow = u.StartRow - count,
                StartColumn = u.StartColumn,
                EndRow = u.EndRow - count,
                EndColumn = u.EndColumn
            };
        }
        else if (u.EndRow >= startRow)
        {
            var overlap = Math.Min(u.EndRow, endRow) - Math.Max(u.StartRow, startRow) + 1;
            UsedRange = new ReportRange
            {
                StartRow = u.StartRow,
                StartColumn = u.StartColumn,
                EndRow = Math.Max(u.StartRow, u.EndRow - overlap),
                EndColumn = u.EndColumn
            };
        }

        RecalculateLayout();
    }

    // Copies the contents (cells and row height) of sourceRow into destinationRow.
    public void CopyRowContent(int sourceRow, int destinationRow)
    {
        if (sourceRow == destinationRow)
        {
            return;
        }

        cells.RemoveAll(c => c.Row == destinationRow);

        var source = GetRowDefinition(sourceRow);
        var destination = GetRowDefinition(destinationRow);
        if (source is not null && destination is not null)
        {
            destination.HeightPoint = source.HeightPoint;
        }

        foreach (var cell in cells.Where(c => c.Row == sourceRow).ToList())
        {
            cells.Add(new ReportCell
            {
                Row = destinationRow,
                Column = cell.Column,
                Value = cell.Value,
                DisplayText = cell.DisplayText,
                Style = cell.Style,
                Merge = cell.Merge
            });
        }

        RecalculateLayout();
    }

    public ReportSheet Clone(string newName)
    {
        var copy = new ReportSheet
        {
            Name = newName,
            UsedRange = UsedRange,
            PageSetup = PageSetup,
            HeaderFooter = HeaderFooter,
            PrintArea = PrintArea,
            ShowGridLines = ShowGridLines
        };

        foreach (var row in rows)
        {
            copy.AddRowDefinition(new ReportRow
            {
                Index = row.Index,
                HeightPoint = row.HeightPoint,
                IsHidden = row.IsHidden,
                OutlineLevel = row.OutlineLevel
            });
        }

        foreach (var col in columns)
        {
            copy.AddColumnDefinition(new ReportColumn
            {
                Index = col.Index,
                WidthPoint = col.WidthPoint,
                IsHidden = col.IsHidden,
                OutlineLevel = col.OutlineLevel,
                OriginalExcelWidth = col.OriginalExcelWidth
            });
        }

        foreach (var mr in mergedRanges)
        {
            copy.AddMergedRange(new ReportMergedRange { Range = mr.Range });
        }

        foreach (var cell in cells)
        {
            copy.AddCell(new ReportCell
            {
                Row = cell.Row,
                Column = cell.Column,
                Value = cell.Value,
                DisplayText = cell.DisplayText,
                Style = cell.Style,
                Merge = cell.Merge
            });
        }

        foreach (var img in images)
        {
            copy.AddImage(new ReportImage
            {
                Name = img.Name,
                FromCellAddress = img.FromCellAddress,
                ToCellAddress = img.ToCellAddress,
                Offset = img.Offset,
                WidthPoint = img.WidthPoint,
                HeightPoint = img.HeightPoint,
                ImageBytes = img.ImageBytes
            });
        }

        foreach (var pb in horizontalPageBreaks)
        {
            copy.AddHorizontalPageBreak(new ReportPageBreak { Index = pb.Index });
        }

        foreach (var pb in verticalPageBreaks)
        {
            copy.AddVerticalPageBreak(new ReportPageBreak { Index = pb.Index });
        }

        copy.RecalculateLayout();
        return copy;
    }

    private void ShiftImageRows(int insertAtRow, int count)
    {
        foreach (var img in images)
        {
            var fromRow = TryParseRow(img.FromCellAddress);
            if (fromRow is not null && fromRow.Value >= insertAtRow)
            {
                img.FromCellAddress = ReplaceRow(img.FromCellAddress, fromRow.Value + count);
            }

            if (img.ToCellAddress is not null)
            {
                var toRow = TryParseRow(img.ToCellAddress);
                if (toRow is not null && toRow.Value >= insertAtRow)
                {
                    img.ToCellAddress = ReplaceRow(img.ToCellAddress, toRow.Value + count);
                }
            }
        }
    }

    private static int? TryParseRow(string address)
    {
        if (String.IsNullOrEmpty(address))
        {
            return null;
        }

        try
        {
            var (row, _) = AddressHelper.ParseAddress(address);
            return row;
        }
        catch (FormatException)
        {
            return null;
        }
    }

    private static string ReplaceRow(string address, int newRow)
    {
        var (_, col) = AddressHelper.ParseAddress(address);
        return AddressHelper.ToAddress(newRow, col);
    }
}

[ExcludeFromCodeCoverage]
internal sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];

    public IReadOnlyList<ReportSheet> Sheets => sheets;

    public ReportMetadata Metadata { get; init; } = new();

    public ReportMeasurementProfile MeasurementProfile { get; init; } = new();

    public void AddSheet(ReportSheet sheet) => sheets.Add(sheet);

    public bool RemoveSheet(ReportSheet sheet) => sheets.Remove(sheet);

    public ReportSheet? FindSheet(string name)
    {
        foreach (var sheet in sheets)
        {
            if (String.Equals(sheet.Name, name, StringComparison.Ordinal))
            {
                return sheet;
            }
        }
        return null;
    }
}
