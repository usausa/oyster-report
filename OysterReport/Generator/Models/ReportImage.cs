namespace OysterReport.Generator.Models;

internal sealed class ReportImage
{
    public ReportImage(
        string name,
        string fromCellAddress,
        string? toCellAddress,
        ReportOffset offset,
        double widthPoint,
        double heightPoint,
        ReadOnlyMemory<byte> imageBytes)
    {
        Name = name;
        FromCellAddress = fromCellAddress;
        ToCellAddress = toCellAddress;
        Offset = offset;
        WidthPoint = widthPoint;
        HeightPoint = heightPoint;
        ImageBytes = imageBytes;
    }

    public string Name { get; }

    public string FromCellAddress { get; }

    public string? ToCellAddress { get; }

    public ReportOffset Offset { get; }

    public double WidthPoint { get; }

    public double HeightPoint { get; }

    public ReadOnlyMemory<byte> ImageBytes { get; }
}
