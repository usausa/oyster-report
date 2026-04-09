namespace OysterReport;

/// <summary>フォント解決時にリゾルバーへ渡す要求情報。</summary>
public sealed record ReportFontRequest
{
    /// <summary>解決対象のフォント名。</summary>
    public string FontName { get; init; } = string.Empty;

    /// <summary>太字指定かどうか。</summary>
    public bool Bold { get; init; }

    /// <summary>斜体指定かどうか。</summary>
    public bool Italic { get; init; }
}
