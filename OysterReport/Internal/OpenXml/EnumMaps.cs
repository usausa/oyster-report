namespace OysterReport.Internal.OpenXml;

using DocumentFormat.OpenXml.Spreadsheet;

internal static class EnumMaps
{
    public static BorderLineStyle ToBorderStyle(BorderStyleValues? style)
    {
        if (style is null)
        {
            return BorderLineStyle.None;
        }

        var raw = style.Value;
        return raw switch
        {
            _ when raw == BorderStyleValues.Thin => BorderLineStyle.Thin,
            _ when raw == BorderStyleValues.Medium => BorderLineStyle.Medium,
            _ when raw == BorderStyleValues.Thick => BorderLineStyle.Thick,
            _ when raw == BorderStyleValues.Double => BorderLineStyle.Double,
            _ when raw == BorderStyleValues.Hair => BorderLineStyle.Hair,
            _ when raw == BorderStyleValues.Dotted => BorderLineStyle.Dotted,
            _ when raw == BorderStyleValues.Dashed => BorderLineStyle.Dashed,
            _ when raw == BorderStyleValues.DashDot => BorderLineStyle.DashDot,
            _ when raw == BorderStyleValues.DashDotDot => BorderLineStyle.DashDotDot,
            _ when raw == BorderStyleValues.MediumDashed => BorderLineStyle.MediumDashed,
            _ when raw == BorderStyleValues.MediumDashDot => BorderLineStyle.MediumDashDot,
            _ when raw == BorderStyleValues.MediumDashDotDot => BorderLineStyle.MediumDashDotDot,
            _ when raw == BorderStyleValues.SlantDashDot => BorderLineStyle.SlantDashDot,
            _ => BorderLineStyle.None
        };
    }

    public static HorizontalAlignment ToHorizontalAlignment(HorizontalAlignmentValues? value)
    {
        if (value is null)
        {
            return HorizontalAlignment.General;
        }

        var raw = value.Value;
        return raw switch
        {
            _ when raw == HorizontalAlignmentValues.Left => HorizontalAlignment.Left,
            _ when raw == HorizontalAlignmentValues.Center => HorizontalAlignment.Center,
            _ when raw == HorizontalAlignmentValues.Right => HorizontalAlignment.Right,
            _ when raw == HorizontalAlignmentValues.Justify => HorizontalAlignment.Justify,
            _ when raw == HorizontalAlignmentValues.CenterContinuous => HorizontalAlignment.CenterContinuous,
            _ when raw == HorizontalAlignmentValues.Distributed => HorizontalAlignment.Distributed,
            _ when raw == HorizontalAlignmentValues.Fill => HorizontalAlignment.Fill,
            _ => HorizontalAlignment.General
        };
    }

    public static VerticalAlignment ToVerticalAlignment(VerticalAlignmentValues? value)
    {
        if (value is null)
        {
            return VerticalAlignment.Bottom;
        }

        var raw = value.Value;
        return raw switch
        {
            _ when raw == VerticalAlignmentValues.Top => VerticalAlignment.Top,
            _ when raw == VerticalAlignmentValues.Center => VerticalAlignment.Center,
            _ when raw == VerticalAlignmentValues.Justify => VerticalAlignment.Justify,
            _ when raw == VerticalAlignmentValues.Distributed => VerticalAlignment.Distributed,
            _ => VerticalAlignment.Bottom
        };
    }

    public static PaperSize ToPaperSize(uint? code)
    {
        if (code is null)
        {
            return PaperSize.A4Paper;
        }

        return Enum.IsDefined(typeof(PaperSize), (int)code.Value)
            ? (PaperSize)code.Value
            : PaperSize.Default;
    }

    public static PageOrientation ToPageOrientation(OrientationValues? value)
    {
        if (value is null)
        {
            return PageOrientation.Default;
        }

        var raw = value.Value;
        return raw switch
        {
            _ when raw == OrientationValues.Portrait => PageOrientation.Portrait,
            _ when raw == OrientationValues.Landscape => PageOrientation.Landscape,
            _ => PageOrientation.Default
        };
    }

    public static FillPattern ToFillPattern(PatternValues? value)
    {
        if (value is null)
        {
            return FillPattern.None;
        }

        var raw = value.Value;
        return raw switch
        {
            _ when raw == PatternValues.None => FillPattern.None,
            _ when raw == PatternValues.Solid => FillPattern.Solid,
            _ when raw == PatternValues.Gray125 => FillPattern.Gray125,
            _ when raw == PatternValues.Gray0625 => FillPattern.Gray0625,
            _ when raw == PatternValues.DarkGray => FillPattern.DarkGray,
            _ when raw == PatternValues.MediumGray => FillPattern.MediumGray,
            _ when raw == PatternValues.LightGray => FillPattern.LightGray,
            _ when raw == PatternValues.DarkHorizontal => FillPattern.DarkHorizontal,
            _ when raw == PatternValues.DarkVertical => FillPattern.DarkVertical,
            _ when raw == PatternValues.DarkDown => FillPattern.DarkDown,
            _ when raw == PatternValues.DarkUp => FillPattern.DarkUp,
            _ when raw == PatternValues.DarkGrid => FillPattern.DarkGrid,
            _ when raw == PatternValues.DarkTrellis => FillPattern.DarkTrellis,
            _ when raw == PatternValues.LightHorizontal => FillPattern.LightHorizontal,
            _ when raw == PatternValues.LightVertical => FillPattern.LightVertical,
            _ when raw == PatternValues.LightDown => FillPattern.LightDown,
            _ when raw == PatternValues.LightUp => FillPattern.LightUp,
            _ when raw == PatternValues.LightGrid => FillPattern.LightGrid,
            _ when raw == PatternValues.LightTrellis => FillPattern.LightTrellis,
            _ => FillPattern.None
        };
    }
}
