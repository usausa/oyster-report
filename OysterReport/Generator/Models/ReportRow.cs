namespace OysterReport.Generator.Models;

internal sealed class ReportRow
{
    public ReportRow(int index, double heightPoint, bool isHidden = false, int outlineLevel = 0)
    {
        Index = index;
        HeightPoint = heightPoint;
        IsHidden = isHidden;
        OutlineLevel = outlineLevel;
    }

    public int Index { get; private set; } // Row number (1-based)

    public double HeightPoint { get; } // Row height (points)

    public double TopPoint { get; private set; } // Top position from the sheet origin (points)

    public bool IsHidden { get; } // Whether the row is hidden

    public int OutlineLevel { get; } // Outline level

    internal ReportRow CloneWithIndex(int index) => new(index, HeightPoint, IsHidden, OutlineLevel);

    internal void SetIndex(int index) => Index = index;

    internal void SetTop(double topPoint) => TopPoint = topPoint;
}
