namespace OysterReport.Reading;

public sealed class ExcelReadOptions
{
    public IReadOnlyList<string>? TargetSheets { get; set; } // 読み込み対象シート名一覧

    public bool IncludeImages { get; set; } = true; // 画像も読み込むか
}
