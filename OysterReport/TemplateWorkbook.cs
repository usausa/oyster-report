namespace OysterReport;

using ClosedXML.Excel;

/// <summary>
/// XLWorkbook を保持し、ワークブック全体のテンプレート操作を提供する。
/// </summary>
public sealed class TemplateWorkbook : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly List<TemplateSheet> _sheets;

    internal TemplateWorkbook(XLWorkbook workbook)
    {
        _workbook = workbook;
        _sheets = workbook.Worksheets.Select(ws => new TemplateSheet(ws)).ToList();
    }

    /// <summary>シート一覧。</summary>
    public IReadOnlyList<TemplateSheet> Sheets => _sheets;

    /// <summary>内部の ClosedXML ワークブック（上級者向け）。</summary>
    [CLSCompliant(false)]
    public IXLWorkbook UnderlyingWorkbook => _workbook;

    /// <summary>名前でシートを取得する。</summary>
    public TemplateSheet GetSheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return _sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"Sheet '{name}' not found.");
    }

    /// <summary>インデックスでシートを取得する (0-based)。</summary>
    public TemplateSheet GetSheet(int index) => _sheets[index];

    /// <summary>テンプレートシートをコピーして新しいシートを作成する。</summary>
    public TemplateSheet CopySheet(string sourceSheetName, string newSheetName)
    {
        ArgumentNullException.ThrowIfNull(sourceSheetName);
        ArgumentNullException.ThrowIfNull(newSheetName);
        var sourceWorksheet = _workbook.Worksheet(sourceSheetName);
        var newWorksheet = sourceWorksheet.CopyTo(newSheetName);
        var newSheet = new TemplateSheet(newWorksheet);
        _sheets.Add(newSheet);
        return newSheet;
    }

    /// <summary>シートを削除する。</summary>
    public void RemoveSheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        var sheet = GetSheet(name);
        _workbook.Worksheet(name).Delete();
        _sheets.Remove(sheet);
    }

    /// <summary>全シートのプレースホルダを一括置換する。</summary>
    public int ReplacePlaceholder(string markerName, string value)
    {
        var count = 0;
        foreach (var sheet in _sheets)
        {
            count += sheet.ReplacePlaceholder(markerName, value);
        }
        return count;
    }

    /// <summary>全シートのプレースホルダを辞書で一括置換する。</summary>
    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        var count = 0;
        foreach (var (key, value) in values)
        {
            count += ReplacePlaceholder(key, value ?? string.Empty);
        }
        return count;
    }

    /// <inheritdoc />
    public void Dispose() => _workbook.Dispose();
}
