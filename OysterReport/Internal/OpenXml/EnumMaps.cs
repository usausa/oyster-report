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
            var v when v == BorderStyleValues.Thin => BorderLineStyle.Thin,
            var v when v == BorderStyleValues.Medium => BorderLineStyle.Medium,
            var v when v == BorderStyleValues.Thick => BorderLineStyle.Thick,
            var v when v == BorderStyleValues.Double => BorderLineStyle.Double,
            var v when v == BorderStyleValues.Hair => BorderLineStyle.Hair,
            var v when v == BorderStyleValues.Dotted => BorderLineStyle.Dotted,
            var v when v == BorderStyleValues.Dashed => BorderLineStyle.Dashed,
            var v when v == BorderStyleValues.DashDot => BorderLineStyle.DashDot,
            var v when v == BorderStyleValues.DashDotDot => BorderLineStyle.DashDotDot,
            var v when v == BorderStyleValues.MediumDashed => BorderLineStyle.MediumDashed,
            var v when v == BorderStyleValues.MediumDashDot => BorderLineStyle.MediumDashDot,
            var v when v == BorderStyleValues.MediumDashDotDot => BorderLineStyle.MediumDashDotDot,
            var v when v == BorderStyleValues.SlantDashDot => BorderLineStyle.SlantDashDot,
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
            var v when v == HorizontalAlignmentValues.Left => HorizontalAlignment.Left,
            var v when v == HorizontalAlignmentValues.Center => HorizontalAlignment.Center,
            var v when v == HorizontalAlignmentValues.Right => HorizontalAlignment.Right,
            var v when v == HorizontalAlignmentValues.Justify => HorizontalAlignment.Justify,
            var v when v == HorizontalAlignmentValues.CenterContinuous => HorizontalAlignment.CenterContinuous,
            var v when v == HorizontalAlignmentValues.Distributed => HorizontalAlignment.Distributed,
            var v when v == HorizontalAlignmentValues.Fill => HorizontalAlignment.Fill,
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
            var v when v == VerticalAlignmentValues.Top => VerticalAlignment.Top,
            var v when v == VerticalAlignmentValues.Center => VerticalAlignment.Center,
            var v when v == VerticalAlignmentValues.Justify => VerticalAlignment.Justify,
            var v when v == VerticalAlignmentValues.Distributed => VerticalAlignment.Distributed,
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
            var v when v == OrientationValues.Portrait => PageOrientation.Portrait,
            var v when v == OrientationValues.Landscape => PageOrientation.Landscape,
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
            var v when v == PatternValues.None => FillPattern.None,
            var v when v == PatternValues.Solid => FillPattern.Solid,
            var v when v == PatternValues.Gray125 => FillPattern.Gray125,
            var v when v == PatternValues.Gray0625 => FillPattern.Gray0625,
            var v when v == PatternValues.DarkGray => FillPattern.DarkGray,
            var v when v == PatternValues.MediumGray => FillPattern.MediumGray,
            var v when v == PatternValues.LightGray => FillPattern.LightGray,
            var v when v == PatternValues.DarkHorizontal => FillPattern.DarkHorizontal,
            var v when v == PatternValues.DarkVertical => FillPattern.DarkVertical,
            var v when v == PatternValues.DarkDown => FillPattern.DarkDown,
            var v when v == PatternValues.DarkUp => FillPattern.DarkUp,
            var v when v == PatternValues.DarkGrid => FillPattern.DarkGrid,
            var v when v == PatternValues.DarkTrellis => FillPattern.DarkTrellis,
            var v when v == PatternValues.LightHorizontal => FillPattern.LightHorizontal,
            var v when v == PatternValues.LightVertical => FillPattern.LightVertical,
            var v when v == PatternValues.LightDown => FillPattern.LightDown,
            var v when v == PatternValues.LightUp => FillPattern.LightUp,
            var v when v == PatternValues.LightGrid => FillPattern.LightGrid,
            var v when v == PatternValues.LightTrellis => FillPattern.LightTrellis,
            _ => FillPattern.None
        };
    }
}
