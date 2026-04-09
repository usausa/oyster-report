namespace OysterReport;

// フォント解決結果。
public sealed record FontInfo
{
    // 描画時に使用するフェース名。
    public string FaceName { get; }

    // 太字を描画側でシミュレーションするかどうか。
    public bool MustSimulateBold { get; set; }

    // 斜体をフォント解決側でシミュレーションするかどうか。
    public bool MustSimulateItalic { get; set; }

    public FontInfo(string faceName)
    {
        FaceName = faceName;
    }
}
