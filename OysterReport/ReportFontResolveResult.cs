namespace OysterReport;

/// <summary>フォント解決結果。</summary>
public sealed record ReportFontResolveResult
{
    /// <summary>
    /// PDFSharp に登録またはシステム検索に使用するフォント名。
    /// </summary>
    public string FontName { get; init; } = string.Empty;

    /// <summary>
    /// 指定した場合はこのバイト列を直接使用し、システムフォント検索を行わない。
    /// </summary>
    public ReadOnlyMemory<byte>? FontData { get; init; }
}
