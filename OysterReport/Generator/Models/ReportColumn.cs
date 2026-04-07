namespace OysterReport.Generator.Models;

internal sealed class ReportColumn
{
    public ReportColumn(int index, double widthPoint, bool isHidden = false, int outlineLevel = 0, double originalExcelWidth = 0)
    {
        Index = index;
        WidthPoint = widthPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
        OriginalExcelWidth = originalExcelWidth;
    }

    public int Index { get; } // Column number (1-based)

    public double WidthPoint { get; } // Column width (points)

    public double LeftPoint { get; private set; } // Left position from the sheet origin (points)

    public bool IsHidden { get; } // Whether the column is hidden

    public int OutlineLevel { get; } // Outline level

    public double OriginalExcelWidth { get; } // Original column width value from Excel

    internal void SetLeft(double leftPoint) => LeftPoint = leftPoint;
}
