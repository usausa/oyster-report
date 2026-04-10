namespace OysterReport.Internal;

using System.Diagnostics.CodeAnalysis;
using System.Globalization;

[ExcludeFromCodeCoverage]
internal readonly record struct ReportOffset
{
    public double X { get; init; }

    public double Y { get; init; }
}

[ExcludeFromCodeCoverage]
internal readonly record struct ReportThickness
{
    public double Left { get; init; }

    public double Top { get; init; }

    public double Right { get; init; }

    public double Bottom { get; init; }
}

[ExcludeFromCodeCoverage]
internal readonly record struct ReportLine
{
    public double X1 { get; init; }

    public double Y1 { get; init; }

    public double X2 { get; init; }

    public double Y2 { get; init; }
}

[ExcludeFromCodeCoverage]
internal readonly record struct ReportRect
{
    public double X { get; init; }

    public double Y { get; init; }

    public double Width { get; init; }

    public double Height { get; init; }

    public double Right => X + Width;

    public double Bottom => Y + Height;

    public ReportRect Deflate(ReportThickness thickness) =>
        new()
        {
            X = X + thickness.Left,
            Y = Y + thickness.Top,
            Width = Math.Max(0, Width - thickness.Left - thickness.Right),
            Height = Math.Max(0, Height - thickness.Top - thickness.Bottom)
        };
}

[ExcludeFromCodeCoverage]
internal readonly record struct ReportRange
{
    public int StartRow { get; init; }

    public int StartColumn { get; init; }

    public int EndRow { get; init; }

    public int EndColumn { get; init; }

    public bool Contains(int row, int column) =>
        row >= StartRow &&
        row <= EndRow &&
        column >= StartColumn &&
        column <= EndColumn;

    public override string ToString()
    {
        var startAddress = AddressHelper.ToAddress(StartRow, StartColumn);
        var endAddress = AddressHelper.ToAddress(EndRow, EndColumn);
        return startAddress == endAddress
            ? startAddress
            : String.Create(CultureInfo.InvariantCulture, $"{startAddress}:{endAddress}");
    }
}
