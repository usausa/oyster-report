namespace OysterReport.Internal;

using System.Globalization;

internal static class ColorHelper
{
    public static string NormalizeHex(string? argb)
    {
        if (String.IsNullOrWhiteSpace(argb))
        {
            return "#00000000";
        }

        var trimmed = argb.AsSpan().Trim();
        if ((trimmed.Length > 0) && (trimmed[0] == '#'))
        {
            var buffer = trimmed.Length <= 64 ? stackalloc char[trimmed.Length] : new char[trimmed.Length];
            for (var i = 0; i < trimmed.Length; i++)
            {
                buffer[i] = Char.ToUpperInvariant(trimmed[i]);
            }

            return new string(buffer);
        }

        var prefixedBuffer = trimmed.Length + 1 <= 64 ? stackalloc char[trimmed.Length + 1] : new char[trimmed.Length + 1];
        prefixedBuffer[0] = '#';
        for (var i = 0; i < trimmed.Length; i++)
        {
            prefixedBuffer[i + 1] = Char.ToUpperInvariant(trimmed[i]);
        }

        return new string(prefixedBuffer);
    }

    public static string ToHex(ArgbColor color)
    {
        Span<char> buffer = stackalloc char[9];
        buffer[0] = '#';
        color.Value.TryFormat(buffer[1..], out _, "X8", CultureInfo.InvariantCulture);
        return new string(buffer);
    }

    public static ArgbColor ApplyTint(ArgbColor color, double tint)
    {
        if (Double.IsNaN(tint))
        {
            return color;
        }

        var clampedTint = Math.Clamp(tint, -1d, 1d);
        if (Math.Abs(clampedTint) < Double.Epsilon)
        {
            return color;
        }

        RgbToHsl(color, out var hue, out var saturation, out var lightness);
        var tintedLightness = clampedTint < 0d
            ? lightness * (1d + clampedTint)
            : lightness + ((1d - lightness) * clampedTint);
        return HslToColor(hue, saturation, tintedLightness, color.A);
    }

    private static void RgbToHsl(ArgbColor color, out double hue, out double saturation, out double lightness)
    {
        var red = color.R / 255d;
        var green = color.G / 255d;
        var blue = color.B / 255d;
        var max = Math.Max(red, Math.Max(green, blue));
        var min = Math.Min(red, Math.Min(green, blue));
        hue = 0d;
        saturation = 0d;
        lightness = (max + min) / 2d;

        if (Math.Abs(max - min) < Double.Epsilon)
        {
            return;
        }

        var delta = max - min;
        saturation = lightness > 0.5d
            ? delta / (2d - max - min)
            : delta / (max + min);

        if (Math.Abs(max - red) < Double.Epsilon)
        {
            hue = ((green - blue) / delta) + (green < blue ? 6d : 0d);
        }
        else if (Math.Abs(max - green) < Double.Epsilon)
        {
            hue = ((blue - red) / delta) + 2d;
        }
        else
        {
            hue = ((red - green) / delta) + 4d;
        }

        hue /= 6d;
    }

    private static ArgbColor HslToColor(double hue, double saturation, double lightness, byte alpha)
    {
        if (saturation <= 0d)
        {
            var value = ToByte(lightness * 255d);
            return new ArgbColor(alpha, value, value, value);
        }

        var q = lightness < 0.5d
            ? lightness * (1d + saturation)
            : lightness + saturation - (lightness * saturation);
        var p = (2d * lightness) - q;
        var red = HueToRgb(p, q, hue + (1d / 3d));
        var green = HueToRgb(p, q, hue);
        var blue = HueToRgb(p, q, hue - (1d / 3d));
        return new ArgbColor(alpha, ToByte(red * 255d), ToByte(green * 255d), ToByte(blue * 255d));
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0d)
        {
            t += 1d;
        }
        if (t > 1d)
        {
            t -= 1d;
        }

        if (t < (1d / 6d))
        {
            return p + ((q - p) * 6d * t);
        }
        if (t < 0.5d)
        {
            return q;
        }
        if (t < (2d / 3d))
        {
            return p + ((q - p) * ((2d / 3d) - t) * 6d);
        }

        return p;
    }

    private static byte ToByte(double value) =>
        (byte)Math.Round(Math.Clamp(value, 0d, 255d), MidpointRounding.AwayFromZero);
}
