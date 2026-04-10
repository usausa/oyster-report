# OysterReport 機能テスト一覧

このドキュメントでは、`OysterReport.Tests/FeatureTests.cs` に実装した機能別テストの一覧と概要を説明する。  
テストは機能ごとにクラスを分けており、`PdfPig` を使って生成した PDF を解析・検証する。  
生成した PDF はすべて `TestOutput/` ディレクトリに保存されるため、目視確認も可能。

---

## テスト対象クラス一覧

| テストクラス | テスト数 | 対象機能 |
|---|---|---|
| `FeatureFontSizeTests` | 5 | フォントサイズ |
| `FeatureFontStyleTests` | 4 | フォントスタイル（太字・斜体） |
| `FeatureFontColorTests` | 4 | フォントカラー |
| `FeatureFillColorTests` | 4 | 背景色（Fill） |
| `FeatureMergedCellTests` | 5 | セル結合 |
| `FeatureImageTests` | 3 | 画像埋め込み |
| `FeatureRowAdditionTests` | 5 | 行の追加 |
| `FeatureTextAlignmentTests` | 5 | テキスト配置 |
| `FeatureBorderTests` | 3 (+6 theory) | 罫線 |
| `FeaturePlaceholderTests` | 4 | プレースホルダー置換 |
| `FeaturePageSetupTests` | 5 | ページ設定 |
| `FeatureHeaderFooterTests` | 3 | ヘッダー・フッター |
| `FeatureMultiSheetTests` | 2 | 複数シート |
| `FeatureHiddenRowColumnTests` | 2 | 非表示行・列 |
| `FeaturePrintAreaTests` | 1 | 印刷範囲 |
| `FeatureCellValueTypeTests` | 3 | セルの値型 |
| `FeatureEmbeddedFontTests` | 5 | 埋め込みフォント（ipaexg.ttf） |
| `FeatureInvoiceTemplateTests` | 2 | 総合テンプレート |

---

## 機能別テスト詳細

### フォントサイズ (`FeatureFontSizeTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainTextWithVariousFontSizes(6pt, "TinyText")` | 6pt の極小フォントでテキストが出力されること |
| `PdfShouldContainTextWithVariousFontSizes(11pt, "NormalText")` | 11pt の標準フォントでテキストが出力されること |
| `PdfShouldContainTextWithVariousFontSizes(18pt, "LargeText")` | 18pt の大フォントでテキストが出力されること |
| `PdfShouldContainTextWithVariousFontSizes(24pt, "HugeText")` | 24pt の極大フォントでテキストが出力されること |
| `PdfShouldRenderMultipleFontSizesOnOnePage` | 8pt・12pt・16pt の混在が 1 ページに出力されること |

**検証方法**: PdfPig でテキスト抽出し、期待文字列が含まれることを確認。

---

### フォントスタイル (`FeatureFontStyleTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainBoldText` | 太字テキストが PDF に出力されること |
| `PdfShouldContainItalicText` | 斜体テキストが PDF に出力されること |
| `PdfShouldContainBoldItalicText` | 太字＋斜体のテキストが PDF に出力されること |
| `PdfShouldRenderMixedNormalBoldItalicOnSamePage` | 通常・太字・斜体が同一ページに共存すること |

**検証方法**: PdfPig でテキスト抽出し、期待文字列が含まれることを確認。

---

### フォントカラー (`FeatureFontColorTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainColoredText("RedText", 255, 0, 0)` | 赤色フォントのテキストが出力されること |
| `PdfShouldContainColoredText("BlueText", 0, 0, 255)` | 青色フォントのテキストが出力されること |
| `PdfShouldContainColoredText("GreenText", 0, 128, 0)` | 緑色フォントのテキストが出力されること |
| `PdfShouldContainThemeColorText` | Excel テーマカラーのフォントが出力されること |

**検証方法**: PdfPig でテキスト抽出し、期待文字列が含まれることを確認。

---

### 背景色 (`FeatureFillColorTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainTextOnColoredBackground("YellowBg", ...)` | 黄色背景セルのテキストが出力されること |
| `PdfShouldContainTextOnColoredBackground("LightBlueBg", ...)` | 水色背景セルのテキストが出力されること |
| `PdfShouldContainTextOnColoredBackground("GrayBg", ...)` | グレー背景セルのテキストが出力されること |
| `PdfShouldRenderMultipleCellsWithDifferentBackgroundColors` | 異なる背景色を持つ複数セルが全て出力されること |
| `PdfShouldRenderThemeBackgroundColor` | Excel テーマカラーの背景色が出力されること |

**検証方法**: PdfPig でテキスト抽出し、期待文字列が含まれることを確認。

---

### セル結合 (`FeatureMergedCellTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainTextInHorizontallyMergedCell` | 横方向結合セル (A1:D1) のテキストが出力されること |
| `PdfShouldContainTextInVerticallyMergedCell` | 縦方向結合セル (A1:A4) のテキストが出力されること |
| `PdfShouldContainTextInRectangularMergedCell` | 矩形結合セル (B2:D4) のテキストが出力されること |
| `PdfShouldRenderMultipleMergedRanges` | 複数の結合セル範囲が共存して出力されること |
| `PdfShouldNotDuplicateTextFromMergedSubCells` | 結合セルのテキストが重複して出力されないこと |

**検証方法**: PdfPig でテキスト抽出・出現回数カウントにより確認。

---

### 画像埋め込み (`FeatureImageTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldEmbedSingleImage` | 単一画像付きシートから PDF が生成されること |
| `PdfShouldEmbedMultipleImages` | 複数画像付きシートから PDF が生成されること |
| `PdfShouldHandleFreeFloatingImage` | FreeFloating 配置の画像が含まれる PDF が生成されること |

**検証方法**: PDF の有効性 (%PDF ヘッダー) とバイトサイズで確認。  
> **NOTE**: PdfPig の `GetImages()` は PDFSharp 生成の XObject 画像を検出できないため、バイトサイズによる間接確認を使用。

---

### 行の追加 (`FeatureRowAdditionTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainAllRowsAddedWithInsertCopyBelow` | `InsertCopyBelow` で追加した全行 (ItemA/B/C) が出力されること |
| `PdfShouldContainAllRowsAddedWithInsertCopyAfter` | `InsertCopyAfter` (Flow A) で追加した全行が出力されること |
| `PdfShouldPreserveStyleAfterRowCopy` | 行コピー後もスタイル（太字・背景色）が維持されること |
| `PdfShouldHandleZeroDetailRows` | 明細行なし（テンプレート削除のみ）でも Header/Footer が出力されること |
| `PdfShouldContainAllRowsFromMultiRowRangeExpansion` | 複数行範囲 (`TemplateRowRange`) の展開で全行が出力されること |

**検証方法**: PdfPig でテキスト抽出し、各行の値が含まれることを確認。

---

### テキスト配置 (`FeatureTextAlignmentTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainHorizontallyAlignedText(Left, ...)` | 左寄せテキストが出力されること |
| `PdfShouldContainHorizontallyAlignedText(Center, ...)` | 中央寄せテキストが出力されること |
| `PdfShouldContainHorizontallyAlignedText(Right, ...)` | 右寄せテキストが出力されること |
| `PdfShouldContainVerticallyAlignedText(Top, ...)` | 上配置テキストが出力されること |
| `PdfShouldContainVerticallyAlignedText(Center, ...)` | 中央配置テキストが出力されること |
| `PdfShouldContainVerticallyAlignedText(Bottom, ...)` | 下配置テキストが出力されること |
| `PdfShouldContainWrappedText` | テキスト折り返し（WrapText）が例外なく出力されること |

**検証方法**: PdfPig でテキスト抽出し、期待文字列が含まれることを確認。

> **NOTE**: WrapText は `XTextFormatter` を使用するため、`XStringFormats.TopLeft` 固定で描画する。

---

### 罫線 (`FeatureBorderTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainCellWithBorder(Thin, ...)` | 細線罫線付きセルが出力されること |
| `PdfShouldContainCellWithBorder(Medium, ...)` | 中線罫線付きセルが出力されること |
| `PdfShouldContainCellWithBorder(Thick, ...)` | 太線罫線付きセルが出力されること |
| `PdfShouldContainCellWithBorder(Double, ...)` | 二重線罫線付きセルが出力されること |
| `PdfShouldContainCellWithBorder(Dashed, ...)` | 破線罫線付きセルが出力されること |
| `PdfShouldContainCellWithBorder(Dotted, ...)` | 点線罫線付きセルが出力されること |
| `PdfShouldRenderColoredBorder` | 赤色罫線付きセルが出力されること |
| `PdfShouldRenderTableWithAllSideBorders` | 全辺罫線の 3×3 テーブルが出力されること |

**検証方法**: PdfPig でテキスト抽出し、期待文字列が含まれることを確認。

---

### プレースホルダー置換 (`FeaturePlaceholderTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainReplacedPlaceholderValue` | 単一プレースホルダー `{{CustomerName}}` の置換値が出力されること |
| `PdfShouldContainAllReplacedPlaceholders` | 複数プレースホルダーの一括置換が全て出力されること |
| `PdfShouldPreservePlaceholderNotReplaced` | 未置換のプレースホルダーを残したまま他の置換が機能すること |
| `PdfShouldContainReplacedPlaceholderInMergedCell` | 結合セル内のプレースホルダー置換値が出力されること |

**検証方法**: PdfPig でテキスト抽出し、置換後の値が含まれることを確認。

---

### ページ設定 (`FeaturePageSetupTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldHaveA4PageDimensions` | A4 縦 (595.28 × 841.89pt) のページサイズであること |
| `PdfShouldHaveA4LandscapePageDimensions` | A4 横 (841.89 × 595.28pt) のページサイズであること |
| `PdfShouldRespectCenterHorizontally` | 水平中央配置設定が有効な状態でテキストが出力されること |
| `PdfShouldGenerateMultiplePagesWhenContentExceedsPageHeight` | 60 行のコンテンツが 1 ページに収まり全テキストが出力されること *(現時点では自動ページ分割未実装)* |
| `PdfShouldApplyManualPageBreak` | 手動ページブレーク設定があっても PDF が正常生成されること *(現時点ではイントラシート改ページ未実装)* |

**検証方法**: PdfPig でページサイズ・ページ数・テキストを確認。

> **NOTE**: `PdfRenderPlanner` は現時点でシートあたり 1 ページプランのみを生成する。自動改ページ・手動改ページのページ分割は未実装。

---

### ヘッダー・フッター (`FeatureHeaderFooterTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldContainHeaderText` | 左ヘッダーが設定されたシートから PDF が生成されること |
| `PdfShouldContainFooterText` | 右フッターが設定されたシートから PDF が生成されること |
| `PdfShouldContainBothHeaderAndFooter` | ヘッダー・フッター両方が設定されたシートから PDF が生成されること |

**検証方法**: PDF の有効性と本文テキストの存在を確認。

---

### 複数シート (`FeatureMultiSheetTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldHaveOnePagePerSheet` | 3 シートのワークブックが 3 ページ以上の PDF になること |
| `PdfShouldContainTextFromAllSheets` | 全シートのテキストが PDF に含まれること |

**検証方法**: PdfPig でページ数・テキストを確認。

---

### 非表示行・列 (`FeatureHiddenRowColumnTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldNotContainHiddenRowText` | 非表示行のテキストが PDF に含まれないこと |
| `PdfShouldNotContainHiddenColumnText` | 非表示列のテキストが PDF に含まれないこと |

**検証方法**: PdfPig でテキスト抽出し、`DoesNotContain` で確認。

---

### 印刷範囲 (`FeaturePrintAreaTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldOnlyContainTextWithinPrintArea` | 印刷範囲外のセルテキストが PDF に含まれないこと |

**検証方法**: PdfPig でテキスト抽出し、範囲内テキストの存在と範囲外テキストの不在を確認。

---

### セルの値型 (`FeatureCellValueTypeTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldRenderNumericCellValue` | 数値セル (12345) の値が PDF に出力されること |
| `PdfShouldRenderDateCellValue` | 日付セル (2025/01/15 形式) の値が PDF に出力されること |
| `PdfShouldRenderFormulaCellValue` | 数式セル (=A1+A2) を含むシートから PDF が生成されること |

**検証方法**: PdfPig でテキスト抽出し、期待文字列が含まれることを確認。

---

### 埋め込みフォント (`FeatureEmbeddedFontTests`)

**前提**: `ipaexg.ttf` がテスト出力ディレクトリに存在すること（プロジェクトで `CopyToOutputDirectory` 設定済み）。

| テスト名 | 内容 |
|---|---|
| `PdfShouldEmbedIpaExGothicFontForJapaneseText` | ゴシック日本語フォントが IPAex として PDF に埋め込まれること |
| `PdfShouldRenderMultipleJapaneseCellsWithEmbeddedFont` | 請求書・合計など複数の日本語セルが出力されること |
| `PdfShouldRenderJapaneseBoldWithEmbeddedFont` | 太字の日本語テキストが埋め込みフォントで出力されること |
| `PdfShouldRenderMixedJapaneseAndAsciiWithEmbeddedFont` | 日本語と ASCII の混在テキストが出力されること |
| `PdfShouldEmbedFontFromIpaExGothicFontPath` | `ipaexg.ttf` ファイルが存在し、PDF が正常に生成されること |

**検証方法**: PdfPig で Letter の FontName が `IPAex` を含むことを確認。

---

### 総合テンプレート (`FeatureInvoiceTemplateTests`)

| テスト名 | 内容 |
|---|---|
| `PdfShouldRenderInvoiceTemplateWithAllFeatures` | 請求書テンプレート（タイトル結合・罫線・プレースホルダー・行追加・背景色）が正常に出力されること |
| `PdfShouldRenderListReportWithStripedRows` | ストライプ行の一覧表 (10 行) が全行出力されること |

**検証方法**: PdfPig でテキスト抽出し、請求書内の主要な値が含まれることを確認。

---

## テスト基盤クラス

### `PdfTestHelper`

| メソッド | 説明 |
|---|---|
| `GeneratePdfAndSave(testName, workbook, resolver?)` | TemplateWorkbook から PDF を生成し `TestOutput/<testName>.pdf` に保存 |
| `GeneratePdfAndSave(testName, stream, resolver?)` | MemoryStream からテンプレートを開いて PDF を生成・保存 |
| `ExtractAllText(pdfBytes)` | 全ページのテキストを結合して返す |
| `ExtractPageText(pdfBytes, pageNumber)` | 指定ページのテキストを返す |
| `GetPageCount(pdfBytes)` | PDF のページ数を返す |
| `GetLetters(pdfBytes, pageNumber)` | 指定ページの `Letter` リストを返す（フォント名・サイズ確認用） |
| `GetPageSize(pdfBytes, pageNumber)` | ページの (幅, 高さ) をポイント単位で返す |
| `GetImageCount(pdfBytes, pageNumber)` | 指定ページの画像数を返す（PdfPig で検出できる場合） |
| `IsValidPdf(pdfBytes)` | `%PDF` ヘッダーで始まる有効な PDF かどうかを返す |
| `IpaExGothicFontPath` | `ipaexg.ttf` のパスを返す |

### `IpaExGothicFontResolver`

ゴシック系日本語フォント（ＭＳ Ｐゴシック・メイリオ等）を `ipaexg.ttf` で解決する `IReportFontResolver` 実装。
`Example/IpaExGothicFontResolver.cs` と同等のものをテストプロジェクト内にコピーして使用する。

---

## 既知の制限事項

| 制限 | 詳細 |
|---|---|
| 自動ページ分割 | `PdfRenderPlanner` はシートあたり 1 ページプランのみ。コンテンツ量による自動改ページは未実装 |
| 手動ページブレーク | `AddHorizontalPageBreak` の設定は読み込まれるが、ページ分割には反映されない |
| PdfPig 画像検出 | PDFSharp 生成の XObject 画像は `page.GetImages()` で検出されない |
| WrapText の垂直配置 | `XTextFormatter` は TopLeft のみ対応のため、WrapText 時は垂直配置が TopLeft 固定 |
