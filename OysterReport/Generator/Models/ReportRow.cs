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

    public int Index { get; }

    public double HeightPoint { get; }

    public double TopPoint { get; private set; }

    public bool IsHidden { get; }

    public int OutlineLevel { get; }

    internal void SetTop(double topPoint) => TopPoint = topPoint;
}
