namespace OysterReport;

public interface IReportFontResolver
{
    /// <summary>
    /// Excel 上のフォント名とスタイル要求から、描画時に使用するフェース名を解決する。
    /// <see langword="null" /> を返した場合は既定のフォント解決にフォールバックする。
    /// </summary>
    FontInfo? ResolveTypeface(string familyName, bool bold, bool italic);

    /// <summary>
    /// フェース名からフォントバイナリを取得する。
    /// <see langword="null" /> を返した場合は既定のフォント取得処理へフォールバックする。
    /// </summary>
    ReadOnlyMemory<byte>? GetFont(string faceName) => null;
}
