// Parses drawings.xml (pictures with anchors + relationships to image parts) into ReportImage.

namespace OysterReport.Prototype;

using System.Globalization;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;

using OysterReport.Internal;

internal static class DrawingLoader
{
    private const double EmuPerInch = 914_400d;
    private const double PointsPerInch = 72d;
    private const double ScreenDpi = 96d;
    private const double EmuToPoint = PointsPerInch / EmuPerInch;
    private const double EmuToPixel = ScreenDpi / EmuPerInch;

    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static IEnumerable<ReportImage> Load(WorksheetPart worksheetPart)
    {
        foreach (var drawingsPart in worksheetPart.DrawingsPart is null ? [] : new[] { worksheetPart.DrawingsPart })
        {
            foreach (var img in LoadFromDrawings(drawingsPart))
            {
                yield return img;
            }
        }
    }

    private static IEnumerable<ReportImage> LoadFromDrawings(DrawingsPart drawingsPart)
    {
        using var stream = drawingsPart.GetStream();
        var doc = XDocument.Load(stream);
        var root = doc.Root;
        if (root is null)
        {
            yield break;
        }

        foreach (var anchor in root.Elements())
        {
            var image = ReadAnchor(anchor, drawingsPart);
            if (image is not null)
            {
                yield return image;
            }
        }
    }

    private static ReportImage? ReadAnchor(XElement anchor, DrawingsPart drawingsPart)
    {
        var localName = anchor.Name.LocalName;
        if (localName is not ("twoCellAnchor" or "oneCellAnchor" or "absoluteAnchor"))
        {
            return null;
        }

        var pic = anchor.Element(Xdr + "pic");
        if (pic is null)
        {
            return null;
        }

        var blip = pic.Descendants(A + "blip").FirstOrDefault();
        var embedId = blip?.Attribute(R + "embed")?.Value;
        if (String.IsNullOrEmpty(embedId))
        {
            return null;
        }

        if (drawingsPart.GetPartById(embedId) is not ImagePart imgPart)
        {
            return null;
        }

        byte[] bytes;
        using (var s = imgPart.GetStream())
        using (var ms = new MemoryStream())
        {
            s.CopyTo(ms);
            bytes = ms.ToArray();
        }

        var name = pic.Element(Xdr + "nvPicPr")?.Element(Xdr + "cNvPr")?.Attribute("name")?.Value ?? string.Empty;

        // Prefer twoCellAnchor (from/to with cell refs + EMU offsets). Fallback to oneCellAnchor's ext size.
        if (localName == "twoCellAnchor")
        {
            var from = anchor.Element(Xdr + "from");
            var to = anchor.Element(Xdr + "to");
            if (from is null || to is null)
            {
                return null;
            }

            var fromCell = EmuAnchorToCellAddress(from);
            var toCell = EmuAnchorToCellAddress(to);
            var (offsetX, offsetY) = ReadCellOffset(from);
            var ext = pic.Element(Xdr + "spPr")?.Element(A + "xfrm")?.Element(A + "ext");
            var widthEmu = ReadLong(ext, "cx");
            var heightEmu = ReadLong(ext, "cy");

            // Mirror ExcelReader.TryGetBottomRightCellAddress: only MoveAndSize keeps ToCellAddress.
            // editAs defaults to "twoCell" (→MoveAndSize); "oneCell"→Move; "absolute"→FreeFloating.
            var editAs = anchor.Attribute("editAs")?.Value;
            var keepToCell = editAs is null or "twoCell";

            return new ReportImage
            {
                Name = name,
                FromCellAddress = fromCell,
                ToCellAddress = keepToCell ? toCell : null,
                Offset = new ReportOffset { X = offsetX * EmuToPixel * PointsPerInch / ScreenDpi, Y = offsetY * EmuToPixel * PointsPerInch / ScreenDpi },
                WidthPoint = widthEmu * EmuToPoint,
                HeightPoint = heightEmu * EmuToPoint,
                ImageBytes = bytes
            };
        }

        if (localName == "oneCellAnchor")
        {
            var from = anchor.Element(Xdr + "from");
            if (from is null)
            {
                return null;
            }

            var fromCell = EmuAnchorToCellAddress(from);
            var (offsetX, offsetY) = ReadCellOffset(from);
            var ext = anchor.Element(Xdr + "ext");
            var widthEmu = ReadLong(ext, "cx");
            var heightEmu = ReadLong(ext, "cy");

            return new ReportImage
            {
                Name = name,
                FromCellAddress = fromCell,
                ToCellAddress = null,
                Offset = new ReportOffset { X = offsetX * EmuToPoint, Y = offsetY * EmuToPoint },
                WidthPoint = widthEmu * EmuToPoint,
                HeightPoint = heightEmu * EmuToPoint,
                ImageBytes = bytes
            };
        }

        return null;
    }

    private static string EmuAnchorToCellAddress(XElement anchorSide)
    {
        var col = Int32.Parse(anchorSide.Element(Xdr + "col")?.Value ?? "0", CultureInfo.InvariantCulture);
        var row = Int32.Parse(anchorSide.Element(Xdr + "row")?.Value ?? "0", CultureInfo.InvariantCulture);
        return AddressHelper.ToAddress(row + 1, col + 1);
    }

    private static (long X, long Y) ReadCellOffset(XElement anchorSide)
    {
        var x = Int64.Parse(anchorSide.Element(Xdr + "colOff")?.Value ?? "0", CultureInfo.InvariantCulture);
        var y = Int64.Parse(anchorSide.Element(Xdr + "rowOff")?.Value ?? "0", CultureInfo.InvariantCulture);
        return (x, y);
    }

    private static long ReadLong(XElement? element, string attr)
    {
        if (element is null)
        {
            return 0L;
        }

        var v = element.Attribute(attr)?.Value;
        return String.IsNullOrEmpty(v) ? 0L : Int64.Parse(v, CultureInfo.InvariantCulture);
    }
}
