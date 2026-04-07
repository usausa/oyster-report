namespace OysterReport.Generator.Models;

using OysterReport.Helpers;

internal sealed class ReportImage
{
    public ReportImage(
        string name,
        ReportAnchorType anchorType,
        string fromCellAddress,
        string? toCellAddress,
        ReportOffset offset,
        double widthPoint,
        double heightPoint,
        ReadOnlyMemory<byte> imageBytes)
    {
        Name = name;
        AnchorType = anchorType;
        FromCellAddress = fromCellAddress;
        ToCellAddress = toCellAddress;
        Offset = offset;
        WidthPoint = widthPoint;
        HeightPoint = heightPoint;
        ImageBytes = imageBytes;

        var (row, _) = AddressHelper.ParseAddress(fromCellAddress);
        FromRow = row;
    }

    public string Name { get; } // Image identifier name

    public ReportAnchorType AnchorType { get; } // Anchor type

    public string FromCellAddress { get; private set; } // Starting cell address

    public string? ToCellAddress { get; private set; } // Ending cell address

    public ReportOffset Offset { get; } // Offset within the starting cell

    public double WidthPoint { get; } // Image width (points)

    public double HeightPoint { get; } // Image height (points)

    public ReadOnlyMemory<byte> ImageBytes { get; } // Raw image data

    internal int FromRow { get; private set; } // Starting row number (1-based)

    internal ReportImage CloneShifted(int rowOffset)
    {
        var (_, fromColumn) = AddressHelper.ParseAddress(FromCellAddress);
        var shiftedFrom = AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        string? shiftedTo = null;
        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = AddressHelper.ParseAddress(ToCellAddress);
            shiftedTo = AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }

        return new ReportImage(Name, AnchorType, shiftedFrom, shiftedTo, Offset, WidthPoint, HeightPoint, ImageBytes);
    }

    internal void ShiftRows(int rowOffset)
    {
        var (_, fromColumn) = AddressHelper.ParseAddress(FromCellAddress);
        FromCellAddress = AddressHelper.ToAddress(FromRow + rowOffset, fromColumn);
        FromRow += rowOffset;

        if (!string.IsNullOrWhiteSpace(ToCellAddress))
        {
            var (toRow, toColumn) = AddressHelper.ParseAddress(ToCellAddress);
            ToCellAddress = AddressHelper.ToAddress(toRow + rowOffset, toColumn);
        }
    }
}
