namespace OysterReport.Internal;

using SkiaSharp;

internal static class FontMetricsHelper
{
    private const double ReferenceScreenDpi = 96d;
    private const double PointsPerInch = 72d;

    /// <summary>
    /// TTF/OTF バイト列から、指定サイズでの最大桁幅を計算する。
    /// 単位: 96 DPI 参照ピクセル (Excel 列幅計算仕様 OOXML §18.3.1.13 に準拠)
    /// </summary>
    /// <returns>計測に失敗した場合は <see langword="null"/>。</returns>
    public static double? MeasureMaxDigitWidth(ReadOnlyMemory<byte> fontData, double fontSizePoints)
    {
        if (fontSizePoints <= 0d || fontData.IsEmpty)
        {
            return null;
        }

        try
        {
            using var skData = SKData.CreateCopy(fontData.ToArray());
            using var typeface = SKTypeface.FromData(skData);
            if (typeface is null)
            {
                return null;
            }

            // Excel の MaxDigitWidth は 96 DPI の論理デバイスコンテキスト上で測定した値。
            // SkiaSharp の SKFont.Size はピクセル単位なので pt × (96/72) に変換する。
            var pixelSize = (float)(fontSizePoints * ReferenceScreenDpi / PointsPerInch);
            using var font = new SKFont(typeface, pixelSize);

            var maxWidth = 0f;
            for (var ch = '0'; ch <= '9'; ch++)
            {
                var width = font.MeasureText(ch.ToString());
                if (width > maxWidth)
                {
                    maxWidth = width;
                }
            }

            return maxWidth > 0f ? maxWidth : null;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException)
        {
            return null;
        }
    }

    /// <summary>
    /// インストール済みフォント名から、指定サイズでの最大桁幅を計算する。
    /// 単位: 96 DPI 参照ピクセル。
    /// </summary>
    public static double? MeasureMaxDigitWidth(string fontFamilyName, double fontSizePoints)
    {
        if (String.IsNullOrWhiteSpace(fontFamilyName) || fontSizePoints <= 0d)
        {
            return null;
        }

        try
        {
            using var typeface = SKTypeface.FromFamilyName(fontFamilyName);
            if (typeface is null)
            {
                return null;
            }

            var pixelSize = (float)(fontSizePoints * ReferenceScreenDpi / PointsPerInch);
            using var font = new SKFont(typeface, pixelSize);

            var maxWidth = 0f;
            for (var ch = '0'; ch <= '9'; ch++)
            {
                var width = font.MeasureText(ch.ToString());
                if (width > maxWidth)
                {
                    maxWidth = width;
                }
            }

            return maxWidth > 0f ? maxWidth : null;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException)
        {
            return null;
        }
    }
}
