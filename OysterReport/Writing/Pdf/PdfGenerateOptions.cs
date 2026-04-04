namespace OysterReport.Writing.Pdf;

public sealed class PdfGenerateOptions
{
    public IReportFontResolver? FontResolver { get; set; } // PDF 描画時に使うフォントリゾルバ

    public bool StrictMode { get; set; } // 未解決要素をエラー扱いする厳格モードか

    public double MinimumReadableScale { get; set; } = 0.5d; // 許容する最小拡大縮小率

    public bool EmbedDocumentMetadata { get; set; } = true; // PDF 文書メタデータを書き込むか

    public bool CompressContentStreams { get; set; } = true; // PDF コンテンツを圧縮するか
}

public interface IReportFontResolver
{
    ReportFontResolveResult Resolve(ReportFontRequest request);
}

public sealed record ReportFontRequest
{
    public string FontName { get; init; } = string.Empty; // 要求フォント名

    public bool Bold { get; init; } // 太字要求フラグ

    public bool Italic { get; init; } // 斜体要求フラグ
}

public sealed record ReportFontResolveResult
{
    public bool IsResolved { get; init; } // 解決成功したか

    public string ResolvedFontName { get; init; } = string.Empty; // 解決後フォント名

    public string? Message { get; init; } // 診断メッセージ
}
