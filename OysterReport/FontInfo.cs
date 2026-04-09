namespace OysterReport;

/// <summary>フォント解決結果。</summary>
public sealed record FontInfo
{
    /// <summary>描画時に使用するフェース名。</summary>
    public string FaceName { get; }

    /// <summary>太字を描画側でシミュレーションするかどうか。</summary>
    public bool MustSimulateBold { get; set; }

    /// <summary>斜体をフォント解決側でシミュレーションするかどうか。</summary>
    public bool MustSimulateItalic { get; set; }

    public FontInfo(string faceName)
    {
        FaceName = faceName;
    }
}
