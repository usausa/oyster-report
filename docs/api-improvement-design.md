# OysterReport API 改善 設計書

## 1. 現状の問題点

### 1.1 現在の公開 API 表面

```text
OysterReportEngine
  Read(filePath) -> ReportWorkbook
  GeneratePdf(workbook, stream)

ReportWorkbook (public)
  Sheets         -> IReadOnlyList<ReportSheet>   <- 不要な構造の公開
  Metadata       -> ReportMetadata               <- 内部情報
  MeasurementProfile -> ReportMeasurementProfile  <- 内部情報
  Diagnostics    -> IReadOnlyList<ReportDiagnostic>
  AddSheet(name) -> ReportSheet                  <- 低水準すぎる
  AddSheet(sheet)

ReportSheet (public)
  Rows, Columns, Cells, MergedRanges, Images    <- 全て不要
  PageSetup, HeaderFooter, PrintArea             <- 全て不要
  ReplacePlaceholder(name, value)                <- 必要
  ReplacePlaceholders(dict)                      <- 必要
  AddRows(RowExpansionRequest)                   <- 必要(改善が必要)
```

### 1.2 本質的な問題

- **過剰な情報公開**: ReportSheet.Rows, .Cells, .Columns 等、使用者がアクセスする必要のない内部構造が全て public
- **モデル型の大量公開**: ReportRow, ReportColumn, ReportCell, ReportCellStyle 等 20 以上の型が公開されている
- **操作の不足**: シートのコピー・削除、テンプレート行の取得、不要行の削除がない
- **行操作の粒度が粗い**: RowExpansionRequest は行番号直指定のバッチ処理のみで、テンプレート行の取得→展開→不要行削除のステップが踏めない

### 1.3 使用者が本来必要とする操作

1. Excel テンプレートの読み込み
2. プレースホルダの値の設定
3. **テンプレート行の取得**（行番号またはマーカーで指定）
4. **テンプレート行からの動的な行追加**（行単位のプレースホルダ対応）
5. **不要行の削除**（行追加によるシート拡張分の調整）
6. テンプレートシートからの動的なシート追加
7. テンプレートシートの削除
8. PDF 生成

---

## 2. 設計方針

### 2.1 基本方針

- **ClosedXML 直接操作**: テンプレート操作（プレースホルダ置換、行展開、シート複製・削除）は ClosedXML の XLWorkbook / IXLWorksheet 上で直接実行する
- **最小公開 API**: ユーザーに見えるのは 7-8 型のみ。内部データ構造やレンダリング情報は一切公開しない
- **内部パイプライン**: PDF 生成は ClosedXML から直接 PDF を描画するのではなく、内部で中間データ構造を経由する 3 段階パイプラインとする
- **共有データ構造**: パイプラインの Phase 1（読み込み）と Phase 2（プラン作成）は同一のデータ構造を使用し、Phase 1 で作成、Phase 2 で更新する

### 2.2 ClosedXML を直接操作対象とする理由

- **シート操作**: IXLWorksheet.CopyTo(), .Delete() でネイティブにシートコピー・削除が可能
- **行操作**: IXLRow.InsertRowsBelow() で書式・結合・罫線を維持した行挿入が可能
- **上級者向け**: UnderlyingWorkbook で ClosedXML の全機能にアクセスでき、ライブラリの想定外の操作にも対応可能
- **メモリ**: テンプレート操作フェーズでは ClosedXML の XLWorkbook のみがメモリに存在し、中間構造は PDF 生成時にのみ作成

### 2.3 内部パイプラインで中間データ構造を経由する理由

- **描画精度**: ClosedXML のオブジェクトをそのまま描画に使うと、列幅の単位変換やスタイル解決のロジックが PdfGenerator に散乱する
- **デバッグ**: 中間データ構造があれば ReportDebugDumper で読み込み結果とレンダリングプランの両方をダンプ出力できる
- **テスト**: 中間データ構造を直接テストでき、ClosedXML のモック不要で描画ロジックを検証できる
- **将来拡張**: 中間データ構造があれば、将来 ClosedXML 以外のリーダーに差し替える際の変更を局所化できる

---

## 3. 公開 API 設計

### 3.1 公開型一覧

| 型名 | 種別 | 責務 |
|---|---|---|
| OysterReportEngine | class | エントリポイント。テンプレートの読み込みと PDF 生成 |
| TemplateWorkbook | class | XLWorkbook を保持し、ワークブック全体の操作を提供 |
| TemplateSheet | class | 1 シートに対するプレースホルダ置換・行操作 |
| SheetRow | class | シート上の 1 行を表し、コピー挿入・プレースホルダ置換・削除を提供 |
| SheetRowRange | class | シート上の連続行範囲を表し、コピー挿入・プレースホルダ置換・削除を提供 |
| PdfGeneratorOption | class | PDF 生成オプション（フォントリゾルバ等） |
| IReportFontResolver | interface | フォント解決の拡張ポイント |
| ReportFontRequest | record | フォント解決要求 |
| ReportFontResolveResult | record | フォント解決結果 |

### 3.2 OysterReportEngine

```csharp
public sealed class OysterReportEngine
{
    /// テンプレート Excel ファイルを読み込む
    TemplateWorkbook Load(string filePath);

    /// テンプレート Excel を Stream から読み込む
    TemplateWorkbook Load(Stream stream);

    /// テンプレートから PDF を生成する
    void GeneratePdf(TemplateWorkbook template, Stream output, PdfGeneratorOption? option = null);
}
```

### 3.3 TemplateWorkbook

```csharp
public sealed class TemplateWorkbook : IDisposable
{
    // -- プロパティ --

    /// シート一覧
    IReadOnlyList<TemplateSheet> Sheets { get; }

    // -- シート取得 --

    /// 名前でシートを取得
    TemplateSheet GetSheet(string name);

    /// インデックスでシートを取得 (0-based)
    TemplateSheet GetSheet(int index);

    // -- シート操作 --

    /// テンプレートシートをコピーして新しいシートを作成
    TemplateSheet CopySheet(string sourceSheetName, string newSheetName);

    /// シートを削除
    void RemoveSheet(string name);

    // -- 全シート一括操作 --

    /// 全シートのプレースホルダを一括置換 (戻り値: 置換件数)
    int ReplacePlaceholder(string markerName, string value);

    /// 全シートのプレースホルダを辞書で一括置換 (戻り値: 置換件数)
    int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values);

    // -- ClosedXML 直接アクセス（上級者向け） --

    /// 内部の ClosedXML ワークブック
    IXLWorkbook UnderlyingWorkbook { get; }

    // -- IDisposable --
    void Dispose(); // XLWorkbook のリソースを解放
}
```

### 3.4 TemplateSheet

```csharp
public sealed class TemplateSheet
{
    // -- プロパティ --

    /// シート名
    string Name { get; }

    // -- プレースホルダ操作 --

    /// マーカー名を指定してプレースホルダを置換 (戻り値: 置換件数)
    int ReplacePlaceholder(string markerName, string value);

    /// 辞書で一括置換 (戻り値: 置換件数)
    int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values);

    // -- 行の取得 --

    /// 行番号で単一行を取得 (1-based)
    SheetRow GetRow(int row);

    /// 行番号で行範囲を取得 (1-based, inclusive)
    SheetRowRange GetRows(int startRow, int endRow);

    /// マーカー名で行を検索して取得
    /// {{markerName}} を含むセルが存在する行を自動検出
    SheetRow FindRow(string markerName);

    /// マーカー名で行範囲を検索して取得
    /// マーカーを含むセルが存在する連続行範囲を自動検出
    SheetRowRange FindRows(string markerName);

    // -- 行の直接操作 --

    /// 指定範囲の行を削除 (1-based, inclusive)
    void DeleteRows(int startRow, int endRow);

    // -- ClosedXML 直接アクセス（上級者向け） --

    /// 内部の ClosedXML ワークシート
    IXLWorksheet UnderlyingWorksheet { get; }
}
```

### 3.5 SheetRow

シート上の 1 行を表す軽量ハンドル。コピー挿入・プレースホルダ置換・削除の 3 操作を提供する。

```csharp
public sealed class SheetRow
{
    // -- プロパティ --

    /// この行の行番号 (1-based)
    int RowNumber { get; }

    // -- コピー挿入 --

    /// この行のコピーを直下に挿入し、挿入された新しい行を返す
    /// - ClosedXML が書式・結合・罫線をコピー
    /// - 後続行は自動的に 1 行下にシフトされる
    SheetRow InsertCopyBelow();

    // -- プレースホルダ操作 --

    /// この行内のプレースホルダを置換 (戻り値: 置換件数)
    int ReplacePlaceholder(string markerName, string value);

    /// この行内のプレースホルダを辞書で一括置換 (戻り値: 置換件数)
    int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values);

    // -- 削除 --

    /// この行を削除する。後続行は自動的に上にシフトされる。
    void Delete();
}
```

### 3.6 SheetRowRange

シート上の連続行範囲を表すハンドル。1 明細が複数行にまたがるテンプレートで使用する。

```csharp
public sealed class SheetRowRange
{
    // -- プロパティ --

    /// 開始行番号 (1-based)
    int StartRow { get; }

    /// 終了行番号 (1-based, inclusive)
    int EndRow { get; }

    /// 行数
    int RowCount { get; }

    // -- コピー挿入 --

    /// この行範囲のコピーを直下に挿入し、挿入された新しい行範囲を返す
    /// - ClosedXML が書式・結合・罫線をコピー
    /// - 後続行は自動的に RowCount 行下にシフトされる
    SheetRowRange InsertCopyBelow();

    // -- プレースホルダ操作 --

    /// この行範囲内のプレースホルダを置換 (戻り値: 置換件数)
    int ReplacePlaceholder(string markerName, string value);

    /// この行範囲内のプレースホルダを辞書で一括置換 (戻り値: 置換件数)
    int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values);

    // -- 削除 --

    /// この行範囲を削除する。後続行は自動的に上にシフトされる。
    void Delete();
}
```

### 3.7 PdfGeneratorOption / IReportFontResolver（既存維持）

```csharp
public sealed class PdfGeneratorOption
{
    IReportFontResolver? FontResolver { get; set; }
    bool EmbedDocumentMetadata { get; set; }
    bool CompressContentStreams { get; set; }
}

public interface IReportFontResolver
{
    ReportFontResolveResult Resolve(ReportFontRequest request);
}

public sealed record ReportFontRequest
{
    string FontName { get; init; }
    bool Bold { get; init; }
    bool Italic { get; init; }
}

public sealed record ReportFontResolveResult
{
    bool IsResolved { get; init; }
    string ResolvedFontName { get; init; }
    string? Message { get; init; }
}
```

---

## 4. 明細行操作の詳細設計

### 4.1 設計方針

旧 `TemplateRowBlock.Expand()` はバッチ操作で「テンプレート行 + 展開済み行」を 1 つのインスタンスに混在させていた。
新設計では **行ごとに独立したインスタンス** (`SheetRow` / `SheetRowRange`) を返し、個別に操作する。

- `InsertCopyBelow()` は **新しい行のインスタンス** を返す（元の行とは別オブジェクト）
- `ReplacePlaceholder()` は **そのインスタンスの行だけ** を対象とする
- `Delete()` は **そのインスタンスの行だけ** を削除する

これにより「テンプレート行」「追加した行」「不要な行」がそれぞれ明確に分離される。

### 4.2 行の取得方法

#### 行番号指定

```csharp
// 単一行（10行目）
SheetRow row = sheet.GetRow(10);

// 複数行範囲（10-12行目）
SheetRowRange range = sheet.GetRows(10, 12);
```

#### マーカー検索

```csharp
// {{detail_item}} を含む行を自動検出（単一行）
SheetRow row = sheet.FindRow("detail_item");

// マーカーを含む連続行範囲を自動検出（複数行）
SheetRowRange range = sheet.FindRows("detail_item");
```

**検出アルゴリズム**:
1. シートの使用範囲内を走査し、 `{{markerName}}` を含むセルを見つける
2. `FindRow`: そのセルの行番号で `SheetRow` を返す
3. `FindRows`: 同一マーカープレフィックスを含む連続行範囲を検出し `SheetRowRange` を返す
4. 見つからない場合は InvalidOperationException をスロー

### 4.3 処理フロー A: テンプレートのコピーを追加していく方式

テンプレート行を起点にコピーを挿入し、各コピーのプレースホルダを個別に置換する。
最後にテンプレート行を削除する。

#### 動作の流れ

```text
[初期状態]
Row 10: {{item}} {{qty}} {{amount}}  <- テンプレート行
Row 11: 合計: {{total}}

[1件目: template.InsertCopyBelow() → row1 (row 11)]
Row 10: {{item}} {{qty}} {{amount}}  <- テンプレート行（そのまま）
Row 11: {{item}} {{qty}} {{amount}}  <- コピーされた行 (row1)
Row 12: 合計: {{total}}              <- 1行下にシフト

[1件目: row1.ReplacePlaceholders({item=商品A, qty=10, amount=1000})]
Row 10: {{item}} {{qty}} {{amount}}  <- テンプレート行（そのまま）
Row 11: 商品A    10       1,000      <- 置換済み (row1)
Row 12: 合計: {{total}}

[2件目: row1.InsertCopyBelow() → row2 (row 12)]
  ※ 直前に処理した行の下にコピーを挿入
Row 10: {{item}} {{qty}} {{amount}}  <- テンプレート行
Row 11: 商品A    10       1,000
Row 12: {{item}} {{qty}} {{amount}}  <- コピーされた行 (row2)
Row 13: 合計: {{total}}

[2件目: row2.ReplacePlaceholders({item=商品B, qty=20, amount=2000})]
Row 10: {{item}} {{qty}} {{amount}}  <- テンプレート行
Row 11: 商品A    10       1,000
Row 12: 商品B    20       2,000      <- 置換済み (row2)
Row 13: 合計: {{total}}

[テンプレート行を削除: template.Delete()]
Row 10: 商品A    10       1,000
Row 11: 商品B    20       2,000
Row 12: 合計: {{total}}
```

#### コード例

```csharp
var sheet = workbook.GetSheet("請求書");

// テンプレート行を取得
var template = sheet.GetRow(10);

// データ件数分コピーして置換
var lastRow = template;
foreach (var item in invoiceItems)
{
    var newRow = lastRow.InsertCopyBelow();
    newRow.ReplacePlaceholders(new Dictionary<string, string?>
    {
        ["item"]   = item.Name,
        ["qty"]    = item.Quantity.ToString(),
        ["amount"] = item.Amount.ToString("#,##0")
    });
    lastRow = newRow;
}

// テンプレート行を削除
template.Delete();
```

### 4.4 処理フロー B: 行番号を進めながら処理する方式

固定行数テンプレート（予備行があらかじめ確保されている）で使用する。
行を追加した分、不要な予備行を同時に削除してシートの行数を一定に保つ。

#### 動作の流れ

```text
[初期状態 - 10行分の予備行が確保されている]
Row 10: {{item}} {{qty}} {{amount}}  <- テンプレート行
Row 11: (予備行)
Row 12: (予備行)
  ...
Row 19: (予備行)
Row 20: 合計: {{total}}

[1件目: row=10]
  GetRow(10) → 行10のインスタンス
  InsertCopyBelow() → 行11にコピー挿入（全体が1行増える）
  DeleteRows(20, 20) → 末尾の予備行を1行削除（全体の行数を戻す）
  行10の ReplacePlaceholders({item=商品A, ...})
  rowNum → 11

[2件目: row=11]
  GetRow(11) → 行11のインスタンス（前回コピーされたテンプレート行）
  InsertCopyBelow() → 行12にコピー挿入
  DeleteRows(20, 20) → 予備行を1行削除
  行11の ReplacePlaceholders({item=商品B, ...})
  rowNum → 12

[3件目: row=12]
  同様...
  rowNum → 13

[ループ後: 行13にはテンプレート行のコピーが残っている]
  DeleteRows(13, 13) → 最後の未使用コピーを削除
  残りの予備行も削除: DeleteRows(13, 19)
```

#### コード例

```csharp
var sheet = workbook.GetSheet("請求書");

int rowNum = 10;  // テンプレート行の行番号
int reservedEndRow = 19;  // 予備行の最終行

foreach (var item in invoiceItems)
{
    var row = sheet.GetRow(rowNum);
    row.InsertCopyBelow();  // 次のイテレーション用にテンプレートをコピー
    sheet.DeleteRows(reservedEndRow, reservedEndRow);  // 予備行を1行削除して行数を維持
    reservedEndRow--;

    row.ReplacePlaceholders(new Dictionary<string, string?>
    {
        ["item"]   = item.Name,
        ["qty"]    = item.Quantity.ToString(),
        ["amount"] = item.Amount.ToString("#,##0")
    });

    rowNum++;
}

// 最後のコピー（未使用テンプレート）と残りの予備行を削除
sheet.DeleteRows(rowNum, reservedEndRow);
```

### 4.5 フロー A と フロー B の比較

| 観点 | フロー A (コピー追加方式) | フロー B (行番号進行方式) |
|---|---|---|
| **適するテンプレート** | 行数が固定されていないテンプレート | 予備行が確保された固定行数テンプレート |
| **シート行数** | データ件数に応じて増える | 常に一定（挿入と削除が対） |
| **不要行の削除** | テンプレート行の Delete() のみ | ループ内で予備行を削除 + ループ後に残りを削除 |
| **コードの明快さ** | シンプル（コピー→置換→テンプレート削除） | やや複雑（行番号管理が必要） |
| **推奨** | 一般的な用途に推奨 | 罫線・レイアウトが厳密に固定されたテンプレートに推奨 |

### 4.6 複数行テンプレートの場合 (SheetRowRange)

1 明細が複数行にまたがるテンプレートでは `SheetRowRange` を使用する。
API は `SheetRow` と同じパターン。

```csharp
// テンプレート: 1明細が3行（品名行・詳細行・区切り行）
var template = sheet.GetRows(10, 12);

var lastRange = template;
foreach (var item in items)
{
    var newRange = lastRange.InsertCopyBelow();
    newRange.ReplacePlaceholders(new Dictionary<string, string?>
    {
        ["item_name"]   = item.Name,
        ["item_detail"] = item.Detail,
        ["item_amount"] = item.Amount.ToString("#,##0")
    });
    lastRange = newRange;
}

// テンプレート行範囲を削除
template.Delete();
```

### 4.7 使用例: 請求書の明細行（完全版）

```csharp
using var engine = new OysterReportEngine();
using var workbook = engine.Load("invoice-template.xlsx");
var sheet = workbook.GetSheet("請求書");

// 1. ヘッダ部のプレースホルダを設定
sheet.ReplacePlaceholders(new Dictionary<string, string?>
{
    ["company_name"] = "株式会社テスト",
    ["invoice_date"] = "2025/07/15",
    ["invoice_no"]   = "INV-2025-001"
});

// 2. テンプレート行を取得
var template = sheet.GetRow(10);

// 3. 明細データでコピー・置換
var lastRow = template;
foreach (var item in invoiceItems)
{
    var newRow = lastRow.InsertCopyBelow();
    newRow.ReplacePlaceholders(new Dictionary<string, string?>
    {
        ["detail_item"]  = item.Name,
        ["detail_qty"]   = item.Quantity.ToString(),
        ["detail_price"] = item.UnitPrice.ToString("#,##0"),
        ["detail_total"] = item.Total.ToString("#,##0")
    });
    lastRow = newRow;
}

// 4. テンプレート行を削除
template.Delete();

// 5. フッタのプレースホルダを設定
sheet.ReplacePlaceholder("grand_total", grandTotal.ToString("#,##0"));

// 6. PDF 生成
using var output = File.Create("invoice.pdf");
engine.GeneratePdf(workbook, output, new PdfGeneratorOption
{
    FontResolver = new JapaneseFontResolver()
});
```

### 4.8 行の削除 (DeleteRows)

テンプレートに予備行が確保されている場合や、展開後に不要な行がある場合に使用する。

#### 実装方針

```text
DeleteRows(startRow, endRow):
  -> IXLWorksheet.Rows(startRow, endRow).Delete()
     - ClosedXML が行を削除し、後続行を自動的に上にシフト
     - 結合範囲が削除対象に含まれる場合は ClosedXML が適切に処理
```

### 4.9 使用例: テンプレートシートの複製と明細展開

```csharp
using var workbook = engine.Load("multi-invoice-template.xlsx");

// 顧客ごとにシートを複製して個別の請求書を作成
foreach (var (customer, index) in customers.Select((c, i) => (c, i)))
{
    var sheetName = $"請求書-{index + 1:D3}";
    var sheet = workbook.CopySheet("テンプレート", sheetName);

    // ヘッダ
    sheet.ReplacePlaceholders(new Dictionary<string, string?>
    {
        ["customer_name"] = customer.Name,
        ["customer_address"] = customer.Address
    });

    // 明細
    var template = sheet.GetRows(10, 12);
    var lastRange = template;
    foreach (var item in customer.Items)
    {
        var newRange = lastRange.InsertCopyBelow();
        newRange.ReplacePlaceholders(new Dictionary<string, string?>
        {
            ["item_name"] = item.Name,
            ["amount"]    = item.Amount.ToString("#,##0")
        });
        lastRange = newRange;
    }
    template.Delete();
}

// テンプレートシートを削除
workbook.RemoveSheet("テンプレート");

// 全シートをまとめて PDF 出力
using var output = File.Create("invoices.pdf");
engine.GeneratePdf(workbook, output);
```

---

## 5. 内部 PDF パイプライン設計

### 5.1 パイプライン全体像

```text
TemplateWorkbook (ClosedXML XLWorkbook)
         |
         |  GeneratePdf() 呼び出し
         v
+-----------------------------------------------------+
|  Phase 1: Read (ClosedXML -> 中間データ作成)          |
|    IXLWorkbook を走査し、RenderWorkbook を構築       |
|    - シート情報、行列定義、セル値・スタイル、        |
|      結合、画像、ページ設定等を読み取り              |
|    - 列幅単位変換、カラー解決等もこのフェーズで実行  |
+--------------------+--------------------------------+
                     |  RenderWorkbook (共有データ構造)
                     v
+-----------------------------------------------------+
|  Phase 2: Plan (中間データを更新・拡張)              |
|    RenderWorkbook のデータを元にレンダリング情報を   |
|    計算し、同じデータ構造に書き戻す                  |
|    - 行位置・列位置の計算 (TopPoint, LeftPoint)      |
|    - セル矩形の計算 (OuterBounds, ContentBounds)     |
|    - ページ分割とページ計画の作成                    |
|    - 画像配置の最終座標計算                          |
|    - ヘッダ/フッタの描画矩形計算                     |
+--------------------+--------------------------------+
                     |  RenderWorkbook (更新済み)
                     v
+-----------------------------------------------------+
|  Phase 3: Draw (PDF 描画)                            |
|    更新済み RenderWorkbook を元に PDFsharp で描画     |
|    - ページごとにセル枠・罫線・テキスト・背景色描画  |
|    - 画像配置                                        |
|    - ヘッダ/フッタ描画                               |
+-----------------------------------------------------+
```

### 5.2 共有データ構造: RenderWorkbook

Phase 1 で作成し、Phase 2 で更新する単一のデータ構造。
現行の ReportWorkbook + PdfRenderPlan を統合したもの。

```text
RenderWorkbook
  +-- Metadata                          <- Phase 1 で作成
  +-- MeasurementProfile                <- Phase 1 で作成
  +-- Diagnostics[]                     <- Phase 1/2 で追加
  |
  +-- Sheets[] : RenderSheet
       +-- Name                         <- Phase 1
       +-- UsedRange                    <- Phase 1
       +-- ShowGridLines                <- Phase 1
       |
       +-- Rows[] : RenderRow
       |    +-- Index, HeightPoint      <- Phase 1 で作成
       |    +-- IsHidden, OutlineLevel  <- Phase 1
       |    +-- TopPoint                <- Phase 2 で計算・書き戻し
       |
       +-- Columns[] : RenderColumn
       |    +-- Index, WidthPoint       <- Phase 1 で作成
       |    +-- IsHidden, OutlineLevel  <- Phase 1
       |    +-- OriginalExcelWidth      <- Phase 1
       |    +-- LeftPoint               <- Phase 2 で計算・書き戻し
       |
       +-- Cells[] : RenderCell
       |    +-- Row, Column, Address    <- Phase 1 で作成
       |    +-- Value, DisplayText      <- Phase 1
       |    +-- Style                   <- Phase 1
       |    +-- Placeholder             <- Phase 1
       |    +-- Merge                   <- Phase 1
       |    +-- OuterBounds             <- Phase 2 で計算・書き戻し
       |    +-- ContentBounds           <- Phase 2 で計算・書き戻し
       |    +-- TextBounds              <- Phase 2 で計算・書き戻し
       |    +-- IsMergedOwner           <- Phase 2 で判定・書き戻し
       |
       +-- MergedRanges[]               <- Phase 1 で作成
       |
       +-- Images[] : RenderImage
       |    +-- Name, AnchorType        <- Phase 1 で作成
       |    +-- FromCell, Offset, Size  <- Phase 1
       |    +-- ImageBytes              <- Phase 1
       |    +-- FinalBounds             <- Phase 2 で計算・書き戻し
       |
       +-- PageSetup                    <- Phase 1 で作成
       +-- HeaderFooter                 <- Phase 1 で作成
       +-- PrintArea                    <- Phase 1 で作成
       +-- PageBreaks[]                 <- Phase 1 で作成
       |
       +-- Pages[] : RenderPage         <- Phase 2 で作成
            +-- PageNumber
            +-- PageBounds
            +-- PrintableBounds
            +-- HeaderFooterInfo
            +-- CellIndices[]           <- Cells 配列へのインデックス参照
```

### 5.3 Phase 1: Read（ClosedXML -> 中間データ作成）

現行 ExcelReader 相当の処理。ClosedXML の XLWorkbook から RenderWorkbook を構築する。

#### 処理内容

```text
ReadWorkbook(IXLWorkbook xlWorkbook) -> RenderWorkbook:
  1. メタデータ・計測プロファイル作成
     - 既定フォント名・サイズ -> MaxDigitWidth 計算
  2. シートごとに:
     a. 行定義の読み取り (Index, Height, IsHidden, OutlineLevel)
     b. 列定義の読み取り (Index, Width -> Point変換, IsHidden)
     c. セルの読み取り
        - 値 (文字列/数値/日付/真偽値/エラー)
        - 表示テキスト (GetFormattedString)
        - スタイル (フォント, 罫線, 背景色, 配置, 書式)
        - カラー解決 (テーマカラー, インデックスカラー -> ARGB)
        - プレースホルダ検出 ({{name}} パターン)
     d. 結合セル範囲の読み取り
     e. 画像の読み取り (アンカー, オフセット, サイズ, バイトデータ)
     f. ページ設定の読み取り (用紙, 余白, 向き, 倍率)
     g. ヘッダ/フッタの読み取り
     h. 印刷範囲・改ページの読み取り
```

ポイント: 現行の ExcelReader のロジックはほぼそのまま流用可能。戻り値の型を RenderWorkbook に変更するだけ。

### 5.4 Phase 2: Plan（中間データの更新・拡張）

現行 PdfRenderPlanner 相当の処理。Phase 1 で作成済みの RenderWorkbook を直接更新する。

#### 処理内容

```text
BuildPlan(RenderWorkbook workbook) -> void:  // 戻り値なし、workbook を in-place 更新
  シートごとに:
    1. ページ境界の計算
       - PageSetup から用紙サイズ・余白・向きを解決
       - PageBounds, PrintableBounds を計算

    2. 行位置・列位置の計算
       - 非表示行を除外し、可視行の TopPoint を累積計算 -> RenderRow に書き戻し
       - 非表示列を除外し、可視列の LeftPoint を累積計算 -> RenderColumn に書き戻し
       - CenterHorizontally / CenterVertically のオフセット適用

    3. セル矩形の計算
       - 各 RenderCell の (Row, Column) から該当 RenderRow/RenderColumn を参照
       - OuterBounds, ContentBounds を計算 -> RenderCell に書き戻し
       - 結合セルの場合は結合範囲全体の矩形を計算
       - テキストオーバーフロー矩形 (TextBounds) を計算

    4. 画像配置の計算
       - FromCell の位置 + Offset から FinalBounds を計算 -> RenderImage に書き戻し

    5. ページ計画の作成
       - 改ページ情報を元にセルをページに分配
       - 各ページの HeaderFooterInfo を計算
       - RenderPage を生成して RenderSheet.Pages に追加
```

ポイント: 現行の PdfRenderPlanner が PdfRenderPlan を新規作成して返していたのに対し、新設計では RenderWorkbook の各フィールドを直接更新する。新たなデータ構造のアロケーションは RenderPage のみ。

### 5.5 Phase 3: Draw（PDF 描画）

現行 PdfGenerator の描画部分。更新済みの RenderWorkbook を受け取り PDF を出力する。

#### 処理内容

```text
DrawPdf(RenderWorkbook workbook, Stream output, PdfGeneratorOption option):
  1. PDFsharp のフォント設定
  2. ページごとに:
     a. PdfPage 作成 (PageBounds のサイズ)
     b. 背景色描画 (各セルの Style.Fill)
     c. 罫線描画 (各セルの Style.Borders)
     d. テキスト描画 (DisplayText, Style.Font, Alignment, TextBounds)
     e. 画像描画 (FinalBounds, ImageBytes)
     f. ヘッダ/フッタ描画 (HeaderFooterInfo)
  3. PDF 保存
```

ポイント: 現行の描画ロジックはほぼそのまま流用。PdfRenderPlan の参照を RenderWorkbook の参照に置き換えるだけ。

### 5.6 Phase 1/2 の統合可能性

Phase 1 と Phase 2 は処理としては分離するが、以下の最適化が可能:

- **行位置の即時計算**: Phase 1 で行定義を読み取った直後に TopPoint を計算し、Phase 2 での再走査を省略
- **列位置の即時計算**: 同上。LeftPoint を Phase 1 の列読み取り時に計算
- **セル矩形の部分計算**: Phase 1 でセルを読み取る際、行列位置が確定していれば Bounds を同時に計算

PDF パイプラインは GeneratePdf() 内部で Phase 1 -> Phase 2 -> Phase 3 を連続実行するため、テンプレート操作後の行番号変更等を考慮する必要はない。そのため Phase 1 時点での先行計算は安全に適用可能。

---

## 6. 内部クラス一覧

### 6.1 中間データ構造（internal）

| クラス名 | 責務 | 現行対応 |
|---|---|---|
| RenderWorkbook | ワークブック全体の中間表現 | ReportWorkbook + PdfRenderPlan |
| RenderSheet | シート単位の中間表現 | ReportSheet + PdfRenderSheetPlan |
| RenderRow | 行定義 + 位置 | ReportRow |
| RenderColumn | 列定義 + 位置 | ReportColumn |
| RenderCell | セル値・スタイル + 描画矩形 | ReportCell + PdfCellRenderInfo |
| RenderImage | 画像データ + 最終配置 | ReportImage + PdfImageRenderInfo |
| RenderPage | ページ計画（Phase 2 で作成） | PdfRenderPagePlan |
| RenderHeaderFooter | ヘッダ/フッタ描画情報 | PdfHeaderFooterRenderInfo |

### 6.2 値型・スタイル型（internal）

| クラス名 | 現行対応 |
|---|---|
| CellValue | ReportCellValue |
| CellStyle | ReportCellStyle |
| CellFont | ReportFont |
| CellFill | ReportFill |
| CellBorders / CellBorder | ReportBorders / ReportBorder |
| CellAlignment | ReportAlignment |
| MergeInfo / MergedRange | ReportMergeInfo / ReportMergedRange |
| PlaceholderText | ReportPlaceholderText |

### 6.3 設定・ジオメトリ型（internal）

| クラス名 | 現行対応 |
|---|---|
| PageSetup | ReportPageSetup |
| HeaderFooterDef | ReportHeaderFooter |
| PrintArea | ReportPrintArea |
| PageBreak | ReportPageBreak |
| Rect | ReportRect |
| Thickness | ReportThickness |
| Line | ReportLine |
| Offset | ReportOffset |
| CellRange | ReportRange |

### 6.4 処理クラス（internal）

| クラス名 | 責務 | 現行対応 |
|---|---|---|
| WorkbookReader | Phase 1: ClosedXML -> RenderWorkbook 変換 | ExcelReader |
| RenderPlanner | Phase 2: RenderWorkbook の in-place 更新 | PdfRenderPlanner |
| PdfWriter | Phase 3: RenderWorkbook -> PDF 描画 | PdfGenerator (描画部分) |
| DumpPayloadFactory | デバッグダンプ用ペイロード作成 | DumpPayloadFactory |

### 6.5 ヘルパー（internal, OysterReport.Helpers）

全て変更なし: AddressHelper, PlaceholderParser, ColumnWidthConverter, FontMeasurementHelper, ColorHelper, PageSizeResolver

---

## 7. 既存型との対応と移行

### 7.1 internal 化する型

現行で public の以下の型はすべて internal に変更し、Render プレフィックスの新型に統合する。

| 現行 public 型 | 処分 |
|---|---|
| ReportWorkbook | -> RenderWorkbook に統合 |
| ReportSheet | -> RenderSheet に統合 |
| ReportRow, ReportColumn, ReportCell | -> RenderRow, RenderColumn, RenderCell に統合 |
| 全スタイル型 | -> CellStyle, CellFont 等に統合 |
| 全ページ設定型 | -> PageSetup, HeaderFooterDef 等に統合 |
| ReportMetadata, ReportMeasurementProfile | -> RenderWorkbook のプロパティとして維持 |
| 全 enum (ReportBorderStyle 等) | internal に変更 |
| ExcelReader | -> WorkbookReader にリネーム |
| PdfGenerator | -> PdfWriter にリネーム（描画部分のみ） |
| PdfRenderPlanner | -> RenderPlanner にリネーム |
| PdfRenderPlan 群 | -> RenderPage, RenderHeaderFooter に統合 |
| ExcelReaderOption | 廃止（Load 後に RemoveSheet で対応） |
| RowExpansionRequest | 廃止（TemplateRowBlock.Expand に置換） |
| ReportDiagnostic, ReportDebugDumper | internal に変更 |
