# OysterReport アーキテクチャ

## 概要

OysterReport は ClosedXML で読み込んだ Excel テンプレートを PDF に変換するライブラリです。
処理は **ExcelReader → PdfRenderPlanner → PdfGenerator** の 3 段パイプラインで構成され、
全体制御は `OysterReportEngine` が担当します。

---

## 処理フロー

```
呼び出し元 (Example / 外部ライブラリ利用者)
    │
    │  1. TemplateWorkbook の作製
    │     new TemplateWorkbook(new XLWorkbook(filePath))
    │
    │  2. プレースホルダ置換・行操作 (任意)
    │     workbook.ReplacePlaceholder(...)
    │     sheet.FindRow(...).InsertCopyBelow()
    │
    ▼
OysterReportEngine.GeneratePdf(template, output)
    │
    │  ─── CreateRenderContext(template) ──────────────────────────────
    │
    │  3. Excel 読み取り
    │     ExcelReader.Read(template.UnderlyingWorkbook)
    │         → ReportWorkbook (全シート・セル・スタイルの中間モデル)
    │
    │  4. PDF レンダリング計画の立案
    │     PdfRenderPlanner.BuildPlan(workbook)
    │         → IReadOnlyList<PdfRenderSheetPlan>
    │           (各セル・画像・ヘッダーの座標・領域を pt 単位で確定)
    │
    │  5. コンテキスト組み立て
    │     new ReportRenderContext { Workbook, SheetPlans, FontResolver, ... }
    │
    │  ──────────────────────────────────────────────────────────────────
    │
    │  6. PDF 書き出し
    │     PdfGenerator.WritePdf(context, output)
    │         → PdfDocument に描画して Stream へ保存
    │
    ▼
  output (Stream) に PDF バイト列が書き込まれる
```

---

## コンポーネント一覧

### ライブラリ (`OysterReport`)

| クラス / ファイル | 種別 | 役割 |
|---|---|---|
| `OysterReportEngine` | `sealed class` | エンドツーエンドの全体制御。PDF 生成設定のプロパティを持つ。 |
| `TemplateWorkbook` | `sealed class` | ClosedXML の `IXLWorkbook` をラップしてプレースホルダ置換・シート操作を提供する。 |
| `TemplateSheet` | `sealed class` | 1 シートのプレースホルダ置換・行検索・行コピーを提供する。 |
| `SheetRow` / `SheetRowRange` | `sealed class` | 行・行範囲への操作 API。 |
| `IReportFontResolver` | `interface` | PDF 生成時のフォント解決戦略。実装を差し替えることで任意フォントを使用できる。 |
| `ReportFontRequest` / `ReportFontResolveResult` | `record` | フォントリゾルバーの入出力 DTO。 |

### ジェネレーター (`OysterReport.Generator`)

| クラス / ファイル | 種別 | 役割 |
|---|---|---|
| `ExcelReader` | `static class` | ClosedXML の `IXLWorkbook` を解析して `ReportWorkbook` を生成する。列幅変換・フォント計測も内包する。 |
| `PdfRenderPlanner` | `static class` | `ReportWorkbook` からセル・画像・ヘッダーの pt 座標を計算して `PdfRenderSheetPlan` 一覧を生成する。用紙サイズ解決も内包する。 |
| `PdfGenerator` | `static class` | `ReportRenderContext` をもとに PDFSharp でドキュメントを描画して `Stream` へ書き出す。 |
| `ReportRenderContext` | `sealed record` | ExcelReader・PdfRenderPlanner の成果物と PDF 生成設定をまとめた持ち回りオブジェクト。 |
| `Models.cs` | 複数の `record` / `class` | `ReportWorkbook` を頂点とする中間モデル群。`ReportSheet`・`ReportCell`・`ReportCellStyle` など。 |
| `Primitives.cs` | `readonly record struct` | `ReportRect`・`ReportRange`・`ReportLine`・`ReportThickness` などの値型プリミティブ。 |
| `PdfRenderPlanner.cs` (上部) | `sealed record` | `PdfRenderSheetPlan`・`PdfRenderPagePlan`・`PdfCellRenderInfo` などのレンダリング計画型。 |
| `PdfRenderingConstants` | `static class` | PDF 描画の調整値（左右セル余白・罫線幅・フォントサイズ既定値など）を一元管理する。 |
| `WindowsInstalledFontResolver` | `class` | Windows 環境でインストール済みフォントを PDFSharp に提供するフォントリゾルバー。 |

### ヘルパー (`OysterReport.Helpers`)

| クラス / ファイル | 役割 |
|---|---|
| `AddressHelper` | セルアドレス文字列 (`"A1"` 等) と行列インデックスの相互変換。 |
| `ColorHelper` | ClosedXML のカラー情報を ARGB16 進文字列へ変換する。テーマカラー解決を含む。 |

### テスト専用 (`OysterReport.Tests`)

| クラス / ファイル | 役割 |
|---|---|
| `DumpPayloadFactory` | `ReportWorkbook` と `PdfRenderSheetPlan` の内容を JSON でダンプするオブジェクトを生成する。 |
| `ReportDebugDumper` | `ReportRenderContext` をもとにワークブックや PDF 準備内容を Stream へ書き出すデバッグ補助クラス。 |
| `WorkbookTestFactory` | テスト用インメモリ Excel ブックを生成するファクトリ。 |

---

## 中間モデル階層

```
ReportWorkbook
├── ReportMetadata          (テンプレート名・ソースパス)
├── ReportMeasurementProfile (列幅計算用フォントメトリクス)
└── ReportSheet[]
    ├── ReportRow[]          (行高・表示/非表示・アウトラインレベル)
    ├── ReportColumn[]       (列幅・表示/非表示)
    ├── ReportCell[]
    │   ├── ReportCellValue  (型別の値)
    │   ├── ReportCellStyle
    │   │   ├── ReportFont
    │   │   ├── ReportFill
    │   │   ├── ReportBorders (Left/Top/Right/Bottom)
    │   │   └── ReportAlignment
    │   └── ReportMergeInfo? (マージ先情報)
    ├── ReportMergedRange[]  (マージセル範囲)
    ├── ReportImage[]        (埋め込み画像)
    ├── ReportPageSetup      (用紙・余白・中央揃え)
    ├── ReportHeaderFooter   (ヘッダー/フッターテキスト)
    ├── ReportPrintArea?     (印刷範囲)
    └── ReportPageBreak[]    (水平/垂直改ページ)
```

---

## 依存関係

```
OysterReportEngine
    ├── TemplateWorkbook / TemplateSheet (Excel 操作層)
    ├── ExcelReader          (ClosedXML → 中間モデル)
    ├── PdfRenderPlanner     (中間モデル → レンダリング計画)
    └── PdfGenerator         (レンダリング計画 → PDFSharp → Stream)

外部ライブラリ:
    ClosedXML  … Excel ファイルの読み書き
    PDFSharp   … PDF ドキュメントの生成・描画
```

---

## 拡張ポイント

| 拡張ポイント | 方法 |
|---|---|
| フォント差し替え | `OysterReportEngine.FontResolver` に `IReportFontResolver` 実装を設定する |
| メタデータ埋め込みの制御 | `OysterReportEngine.EmbedDocumentMetadata` を切り替える |
| PDF 圧縮の制御 | `OysterReportEngine.CompressContentStreams` を切り替える |
| 描画調整値の変更 | `PdfRenderingConstants` の定数を変更する |
