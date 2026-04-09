namespace OysterReport;

using ClosedXML.Excel;

// Excel ファイルを保持し、ワークブック全体のテンプレート操作を提供する。
public sealed class TemplateWorkbook : IDisposable
{
    private readonly XLWorkbook workbook;
    private readonly List<TemplateSheet> sheets;

    // ファイルパスから Excel ファイルを読み込む。
    public TemplateWorkbook(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);

        workbook = new XLWorkbook(filePath);
        sheets = workbook.Worksheets.Select(ws => new TemplateSheet(ws)).ToList();
    }

    // ストリームから Excel ファイルを読み込む。
    public TemplateWorkbook(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        workbook = new XLWorkbook(stream);
        sheets = workbook.Worksheets.Select(ws => new TemplateSheet(ws)).ToList();
    }

    // 内部の ClosedXML ワークブック。
    internal IXLWorkbook UnderlyingWorkbook => workbook;

    // シート一覧。
    public IReadOnlyList<TemplateSheet> Sheets => sheets;

    // 名前でシートを取得する。
    public TemplateSheet GetSheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return sheets.FirstOrDefault(s => String.Equals(s.Name, name, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"Sheet '{name}' not found.");
    }

    // インデックスでシートを取得する (0-based)。
    public TemplateSheet GetSheet(int index) => sheets[index];

    // テンプレートシートをコピーして新しいシートを作成する。
    public TemplateSheet CopySheet(string sourceSheetName, string newSheetName)
    {
        ArgumentNullException.ThrowIfNull(sourceSheetName);
        ArgumentNullException.ThrowIfNull(newSheetName);
        var sourceWorksheet = workbook.Worksheet(sourceSheetName);
        var newWorksheet = sourceWorksheet.CopyTo(newSheetName);
        var newSheet = new TemplateSheet(newWorksheet);
        sheets.Add(newSheet);
        return newSheet;
    }

    // シートを削除する。
    public void RemoveSheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        var sheet = GetSheet(name);
        workbook.Worksheet(name).Delete();
        sheets.Remove(sheet);
    }

    // 全シートのプレースホルダを一括置換する。
    public int ReplacePlaceholder(string markerName, string value)
    {
        var count = 0;
        foreach (var sheet in sheets)
        {
            count += sheet.ReplacePlaceholder(markerName, value);
        }
        return count;
    }

    // 全シートのプレースホルダを辞書で一括置換する。
    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values)
    {
        var count = 0;
        foreach (var (key, value) in values)
        {
            count += ReplacePlaceholder(key, value ?? string.Empty);
        }
        return count;
    }

    public void Dispose() => workbook.Dispose();
}
