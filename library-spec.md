# OysterReport ライブラリ仕様

## クラス一覧

### 入口 API

- `OysterReportEngine`
  - ライブラリの高水準エントリポイント
- `ExcelReader`
  - Excel テンプレートを読み込み、中間データ構造を生成する
- `PdfGenerator`
  - 内部で 2 段階処理を行い、PDF を生成する

### Excel 主体の中間データ構造

- `ReportWorkbook`
  - 読み込み結果全体を保持し、シート追加メソッドを持つルートモデル
- `ReportMetadata`
  - ワークブック全体の作成元やテンプレート情報を保持する
- `ReportSheet`
  - シート単位のソースモデルで、プレースホルダ置換や行追加メソッドを持つ
- `ReportRow`
  - 行定義と行位置を保持するソースモデル
- `ReportColumn`
  - 列定義と列位置を保持するソースモデル
- `ReportCell`
  - セルの値、表示文字列、スタイル、矩形を保持するソースモデル
- `ReportPlaceholderText`
  - Excel 上の特殊値から抽出したプレースホルダ情報
- `ReportMergedRange`
  - 結合セル範囲を表すソースモデル
- `ReportPageSetup`
  - 用紙サイズ、余白、倍率など印刷設定を表す
- `ReportHeaderFooter`
  - ヘッダ、フッタの原文定義を表す
- `ReportPrintArea`
  - 印刷範囲を表す
- `ReportPageBreak`
  - 手動改ページを表す
- `ReportRange`
  - Excel 上の範囲指定を表す
- `ReportRect`
  - point 基準の矩形を表す共通型
- `ReportThickness`
  - 余白やパディングを表す共通型
- `ReportLine`
  - 線分を表す共通型
- `ReportImage`
  - 画像アンカー、オフセット、元データ参照を表す
- `ReportMeasurementProfile`
  - 列幅換算、DPI、既定フォントなど環境差吸収の前提を保持する
- `ReportDiagnostic`
  - 読み込み、変換、描画で発生した警告や差異を表す

### 帳票変換要求

- `RowExpansionRequest`
  - 行追加の対象シート、テンプレート範囲、追加データを指定する

### 内部レンダリング用 (`internal`)

- `PdfRenderPlanner`
  - 中間データ構造から PDF 用レンダリング情報を生成する内部クラス
- `PdfRenderPlan`
  - PDF 描画直前のレンダリング情報全体を表す内部クラス
- `PdfRenderSheetPlan`
  - シート単位のレンダリング情報を表す内部クラス
- `PdfRenderPagePlan`
  - ページ単位のレンダリング情報を表す内部クラス
- `PdfCellRenderInfo`
  - セルの描画矩形やクリップ結果を保持する内部クラス
- `PdfBorderRenderInfo`
  - 共有辺競合解決後の罫線セグメントを保持する内部クラス
- `PdfImageRenderInfo`
  - 描画位置確定後の画像レンダリング情報を保持する内部クラス
- `PdfHeaderFooterRenderInfo`
  - 各ページに描画するヘッダ、フッタ情報を保持する内部クラス

### デバッグと運用補助

- `ReportDebugDumper`
  - 中間データ構造や PDF 生成前準備結果をデバッグ用にダンプ出力する
- `ReportDumpFormat`
  - ダンプ出力形式を表す

### オプションと設定

- `ExcelReadOptions`
  - Excel 読み込み時の対象範囲や画像読込条件を指定する
- `PdfGenerateOptions`
  - PDF 生成全体の条件を指定する

## 1. 目的

OysterReport は、Excel ファイルを帳票テンプレートとして読み込み、Excel の見た目を可能な範囲で再現した PDF を生成する .NET ライブラリとする。

Excel はレイアウト編集の正とし、アプリケーション側はそのテンプレートを読み込んだ中間データ構造に対して、プレースホルダ置換や明細行追加のような帳票生成に必要な限定操作だけを行う。最終的な PDF は、その中間データ構造を元に `PdfGenerator` が内部で 2 段階処理を行って生成する。

### 1.1 この仕様で実現したいこと

- ClosedXML で Excel ファイルを読み取る
- Excel の内容を、中間データ構造として保持する
- 中間データ構造には、印字に必要なレイアウト、色、フォント、修飾、印刷設定を持たせる
- Excel はレイアウト編集の主体とし、レイアウト変更 API は基本的に提供しない
- 中間データ構造は、プレースホルダ置換や行追加のような帳票処理に必要な範囲で編集可能にする
- PDFsharp / PDFsharp-MigraDoc を使って PDF を生成する

### 1.2 仕様のポイント

- 余白、各セルのサイズ、罫線、結合、改ページを含め、Excel レイアウトを可能な限り再現する
- フォント情報もなるべく再現し、フォールバックや未解決時の影響が分かるようにする
- 中間データ構造は Excel 主体のソース情報に徹し、PDF 描画用の計算結果は `PdfGenerator` の内部情報として分離する
- 実装が複雑になる項目は、理由と段階導入方針を仕様に明記する
- 他ライブラリの方が優れている箇所は、その方針に近づける
- spec だけで「何を作りたいか」と「どういう設計方針で進めるか」が分かるようにする

本ライブラリは以下の 4 段階で構成する。

1. ClosedXML で Excel を読み取り、中間データ構造へ変換する。
2. `ReportWorkbook` / `ReportSheet` のメソッドで、プレースホルダ置換や行追加など限定された帳票変換を行う。
3. `PdfGenerator` の内部で、PDF 描画用の internal レンダリング情報を生成する。
4. `PdfGenerator` の内部で、internal レンダリング情報と元の中間データ構造を使って PDFsharp / PDFsharp-MigraDoc で PDF を生成する。

Excel をそのまま PDF に変換するだけでなく、テンプレートのセルを値に置換したり、データ件数に応じてテンプレート行を増やしたりできることを主要要件とする。

## 2. 設計方針

### 2.1 基本方針

- Excel はテンプレート入力フォーマットとして扱う。
- PDF 出力の直前に直接 Excel を参照するのではなく、中間データ構造を唯一の描画元とする。
- 中間データ構造は Excel 側の情報を主体とし、PDF 描画専用の計算結果は保持しない。
- 中間データ構造のプロパティは、外部から不用意に変更できない読み取り専用設計を基本とする。
- PDF 出力時は Excel の印刷設定を可能な限り尊重する。
- レイアウトの忠実な再現という観点で他ライブラリの設計が優れている箇所は、その方針へ寄せる。
- プログラムからの変更は `ReportWorkbook.AddSheet`、`ReportSheet.ReplacePlaceholder`、`ReportSheet.AddRows` のような制御されたメソッドに限定する。
- PDF 出力は「中間データ構造からレンダリング情報を生成するフェーズ」と「レンダリング情報から PDF を描画するフェーズ」の 2 段階で行う。
- 非表示行、非表示列は中間データ構造には保持するが、v1 の PDF レンダリング対象からは除外する。
- 完全再現を目標にするのではなく、「帳票用途で破綻しない再現性」を優先する。

### 2.2 v1 のスコープ

v1 で必須とする対象は以下。

- セルの文字列、数値、日付、真偽値の表示
- 行高、列幅、セル位置、結合セル
- 罫線、背景色、文字色、フォント、文字配置、折り返し
- ワークシートの使用範囲
- 余白、用紙サイズ、向き、印刷範囲、改ページ
- ヘッダ、フッタのテキスト
- セルに設定した特殊値プレースホルダの置換
- 画像の貼り付け位置情報の保持と PDF への描画

v1 では原則として非対象または限定対応とする。

- Excel 数式の再計算
- グラフ、SmartArt、フォーム部品、コメント
- 条件付き書式の再評価
- VBA / マクロ
- Excel 独自の高度な描画効果の完全再現

### 2.3 環境差対応方針

OysterReport は、実行環境による差異が出やすい箇所を明示的に扱う。

- フォント
  - OS ごとのインストールフォント差異を前提に、`IReportFontResolver` とフォールバック診断で吸収する
- 列幅、文字幅
  - `ReportMeasurementProfile` により、DPI、既定フォント、列幅補正係数を明示する
- 画像
  - 画像デコーダや埋め込み形式差異に備えて、読み込み不能時は診断を出す
- ロケール
  - 数値、日付、通貨の表示は Excel に保存された表示文字列を優先し、実行環境のカルチャ差異を最小化する
- 改ページ
- フォント計測差異によりページ数がずれる可能性があるため、レンダリング情報とダンプ出力で検証可能にする

環境差の影響を受けやすい箇所は「隠す」のではなく、設定と診断で可視化する方針を採る。

## 未サポート機能一覧

以下は v1 で未サポート、または限定対応とする機能である。

### Excel 機能

- 数式の再計算
- 条件付き書式の再評価
- VBA / マクロ
- コメント
- データ検証 UI
- ピボットテーブルの再解釈

### 描画要素

- グラフ
- SmartArt
- フォーム部品
- 高度な図形描画
- Excel 独自の描画効果の完全再現

### レイアウト再現の限定対応

他ライブラリ欄では、PDF / 印刷観点での扱いのみを記載する。`excelize` は PDF レンダラではないため、主に Excel 構造を保持できるかどうかの観点で記載する。

- 斜め罫線
  - 理由: 共有辺セグメント中心の罫線モデルとは別管理が必要で、PDFsharp 上でも Excel 線種との一致が難しいため
  - 他ライブラリ:
    - Excel.Report.PDF: 共有辺優先ロジックはあるが、斜め罫線の完全再現は確認できていない
    - ReoGrid: 印刷範囲と改ページ管理は強いが、PDF 帳票向けの斜め罫線再現は主対象ではない
    - excelize: Open XML 上の形状は扱えるが、PDF 描画機能は持たない
    - DioDocs for Excel: 高忠実度製品として対応期待が高い
- 一部の特殊罫線種別
  - 理由: Hairline、複雑な二重線、特殊破線は PDFsharp で近似描画になりやすいため
  - 他ライブラリ:
    - Excel.Report.PDF: 優先順位付き罫線描画を持つが、線種は一部近似になる
    - ReoGrid: 境界管理は明確だが、Excel 互換の特殊線種再現を主眼にはしていない
    - excelize: 線種情報の保持はできるが、PDF 描画までは行わない
    - DioDocs for Excel: 商用製品として広い再現範囲が期待できる
- 隣接セルへの文字オーバーフロー完全再現
  - 理由: 隣接セルの空状態、結合状態、罫線、回転文字との相互作用が多く、ページ分割とも衝突しやすいため
  - 他ライブラリ:
    - Excel.Report.PDF: 基本はセル描画矩形中心で、Excel 互換の空セル越え再現は限定的
    - ReoGrid: 印刷用テキスト境界を独立管理できるが、Excel の空セル越え互換そのものは別問題
    - excelize: AutoFit は近似計算で、PDF 上の文字オーバーフロー再現機能は持たない
    - DioDocs for Excel: 高忠実度再現の対象になりやすい
- 巨大な結合セルがページ境界をまたぐケースの完全再現
  - 理由: 分断禁止を徹底すると極端な縮小や空白ページが発生しうる一方、分断すると Excel と見た目が変わるため
  - 他ライブラリ:
    - Excel.Report.PDF: owner cell 中心の描画で比較的単純に処理する
    - ReoGrid: 印刷範囲と改ページの構造は強いが、巨大結合セルの Excel 帳票互換が主目的ではない
    - excelize: 結合範囲の正規化は強いが、PDF 印刷時の分断制御は対象外
    - DioDocs for Excel: 製品思想としては対応期待が高い
- 画像の高度なトリミングや回転
  - 理由: Open XML の画像変換情報をすべて解釈する必要があり、帳票用途の優先度に対して v1 の実装コストが高いため
  - 他ライブラリ:
    - Excel.Report.PDF: 画像差し込みはあるが、高度な変形までを中心機能にはしていない
    - ReoGrid: 印刷対象として画像は扱えるが、Excel 画像変換の互換再現は主対象ではない
    - excelize: 画像アンカーや図形追加 API はあるが、PDF 描画は行わない
    - DioDocs for Excel: 高忠実度製品として対応期待が高い
- ヘッダ、フッタ画像
  - 理由: VML ベースの画像参照解釈と本文領域とのレイアウト競合解決が必要なため
  - 他ライブラリ:
    - Excel.Report.PDF: ソース上ではヘッダ、フッタ余白は扱うが、画像付きヘッダ、フッタの中心機能は確認できていない
    - ReoGrid: 印刷ページ管理はあるが、Excel 互換のヘッダ、フッタ画像は主対象ではない
    - excelize: `SetHeaderFooter` と `AddHeaderFooterImage` を持ち、Excel 構造としては扱える
    - DioDocs for Excel: 高忠実度製品として対応期待が高い

未サポート機能は、可能な限り診断へ記録し、黙って欠落させない。

## 要検討項目

この章は、各章に分散している「検討が必要な項目」のみを横断的にまとめたものである。各項目について、主な影響と取り得る方針を併記する。各方針名の括弧内には、その方向性に近い設計や実装を採っている参照ライブラリを記載する。

### 1. 列幅換算と文字幅計測

主な影響:

- Excel 実機と PDF 出力で列幅や改行位置がずれる
- OS、フォント、DPI の違いでページ数が変わる
- 日本語を含む帳票で見た目差が大きくなりやすい

取り得る方針:

- 方針1(Excel.Report.PDF、excelize)
  - Excel.Report.PDF や excelize は、列幅換算式を専用関数に閉じ込め、近似式ベースで処理する
  - 実装は比較的単純で安定する
  - 環境差の吸収力は低く、Excel 実機との差が残る可能性がある
- 方針2(ReoGrid)
  - ReoGrid は、レンダラや実測結果を強く使い、テキスト計測結果をレイアウトへ反映する
  - 画面や印刷での見た目に寄せやすい
  - 実行環境依存が強くなり、再現性管理が難しくなる
- 方針3(OysterReport 候補)
  - 近似式を基本にしつつ、`ReportMeasurementProfile` で補正可能にする
  - 実装複雑度と調整可能性のバランスが良い
  - 利用者または検証側にプロファイル調整の判断が必要になる

現在の方針:

- v1 は `ReportMeasurementProfile` による補正可能方式を採る
- 差異は診断とダンプで追跡可能にする

### 2. フォント解決とフォールバック

主な影響:

- フォント未解決時に文字幅、改行、ページ数が変わる
- 一部文字が豆腐や欠字になる可能性がある
- 日本語帳票で見た目差が大きくなりやすい

取り得る方針:

- 方針1(Excel.Report.PDF)
  - Excel.Report.PDF は、`IFontResolver` 相当で使用フォントを明示的に解決する
  - 必要に応じてフォールバックを実装しやすい
  - 実装者がフォント資産を管理する必要がある
- 方針2(ReoGrid)
  - ReoGrid は、実行環境の描画基盤に寄せてフォント計測と表示を行う
  - 実環境で自然な見え方を得やすい
  - 環境が変わると見た目差異が出やすい
- 方針3(OysterReport 候補)
  - 明示的なリゾルバとフォールバック診断を持ち、必要に応じて厳格モードを用意する
  - 実運用と忠実度の両立を狙いやすい
  - ルール設計と運用方針の整備が必要になる

現在の方針:

- `IReportFontResolver` とフォールバック診断を前提とする
- 必要に応じて厳格モードで未解決フォントをエラー化できる余地を残す

### 3. 隣接セルへの文字オーバーフロー

主な影響:

- Excel では見えていた文字が PDF では切れる
- 空セルをまたぐ見た目が再現されない
- 帳票の見栄えに直接影響する

取り得る方針:

- 方針1(Excel.Report.PDF、excelize)
  - Excel.Report.PDF や excelize は、セル矩形または描画矩形内に収める前提で処理する
  - 実装は単純で予測しやすい
  - Excel の空セル越え表示は再現しにくい
- 方針2(ReoGrid)
  - ReoGrid は、テキスト矩形を独立管理し、印刷レイアウトで調整余地を持たせる
  - 将来的な忠実再現へ発展させやすい
  - 完全な Excel 互換には追加判定が必要
- 方針3(OysterReport 候補)
  - v1 はクリップ前提を基本にしつつ、`TextBounds` とクリップ情報を保持する
  - 後続段階で空セル越え表示を導入しやすい
  - 初期段階では見た目差が残る

現在の方針:

- v1 は `TextBounds` とクリップ状態を保持し、限定対応とする
- 完全再現は後続段階の検討項目とする

### 4. 特殊罫線の再現

主な影響:

- 二重線、斜め罫線、Hairline などで Excel と見た目差が出る
- 共有辺の優先順位が崩れると罫線が濃くなったり欠けたりする

取り得る方針:

- 方針1(Excel.Report.PDF)
  - Excel.Report.PDF は、共有辺に優先順位を与え、二重描画を避けながら描く
  - 表形式帳票の忠実度が上がる
  - 特殊線種は近似になりやすい
- 方針2(ReoGrid)
  - ReoGrid は、境界線をセル内容とは別構造で管理し、結合や行列変更時に再構成する
  - 境界整合は取りやすい
  - セグメント管理の実装が重くなる
- 方針3(OysterReport 候補)
  - 共有辺優先順位と境界セグメント正規化を組み合わせる
  - 帳票で重要な罫線忠実度を優先できる
  - 斜め罫線や特殊線種は段階導入になる

現在の方針:

- Excel.Report.PDF に近い共有辺優先順位解決を採用する
- v1 は外周、共有辺、二重線近似を優先し、特殊線種は段階導入とする

### 5. 結合セルと改ページの競合

主な影響:

- ページ境界で結合セルが分断される
- 極端な縮小や不自然な空白ページが発生する
- テキスト配置や画像配置が崩れる

取り得る方針:

- 方針1(ReoGrid、excelize)
  - ReoGrid や excelize は、結合セル範囲の整合性を最優先し、開始点、終了点、矩形を常に同期させる
  - レイアウト破綻を防ぎやすい
  - ページ境界との競合時は別途判断が必要になる
- 方針2(Excel.Report.PDF)
  - Excel.Report.PDF は、結合セルの先頭セルを中心に描画を管理する
  - 描画処理は比較的単純になる
  - 結合セルの分断回避ロジックは弱くなりやすい
- 方針3(OysterReport 候補)
  - 構造整合は ReoGrid、excelize 寄りに持ち、描画は owner cell 中心で最適化する
  - ページ分断は原則禁止とし、回避不能時のみ近似処理と診断を行う
  - 判定ロジックがやや複雑になる

現在の方針:

- v1 は分断しないことを第一優先とする
- 回避不能時のみ診断付きで近似処理を認める

### 6. 手動改ページと拡大縮小指定の競合

主な影響:

- Excel の印刷設定と異なるページ数になる
- 横 1 ページ指定と手動改ページが衝突した際に期待通りにならない
- 可読性優先か設定忠実度優先かで結果が変わる

取り得る方針:

- 方針1(Excel.Report.PDF)
  - Excel.Report.PDF は、余白、倍率、横幅合わせを同じスケール決定で処理し、Excel 設定を強く尊重する
  - テンプレート忠実度は高い
  - 可読性が下がる場合がある
- 方針2(ReoGrid)
  - ReoGrid は、印刷範囲、手動改ページ、自動改ページを分離し、競合を構造的に扱う
  - 判断理由を整理しやすい
  - Excel と完全に同じ優先順位にするには追加定義が必要
- 方針3(OysterReport 候補)
  - Excel の明示設定優先を基本としつつ、競合理由を `PdfRenderPagePlan` と診断へ残す
  - 忠実度と調査性の両立を狙える
  - 実装上は解決ルールの明文化が必要になる

現在の方針:

- v1 は Excel の明示設定優先を基本とする
- 可読性が極端に下がる場合は診断を出す

### 7. 画像、図形、チャートのサポート範囲

主な影響:

- 帳票内のロゴや署名は重要度が高い一方、図形やチャートまで含めると実装量が急増する
- 図形未対応だとテンプレート互換の印象が下がる

取り得る方針:

- 方針1(Excel.Report.PDF、excelize)
  - Excel.Report.PDF や excelize は、まず画像を優先し、アンカーとオフセットを正しく再現する
  - 帳票用途では効果が大きい
  - 図形、チャート中心のテンプレートには不足が残る
- 方針2(DioDocs for Excel)
  - DioDocs for Excel は、図形やチャートも含めてテンプレート再現対象に広げる
  - 製品としての忠実度は高い
  - v1 のスコープとしては重くなる
- 方針3(OysterReport 候補)
  - v1 は画像優先としつつ、中間構造は図形拡張可能にしておく
  - 初期コストを抑えながら拡張余地を残せる
  - v1 の時点では図形テンプレートで不足が出る

現在の方針:

- v1 は画像を優先し、図形とチャートは将来対応とする
- 中間構造は将来拡張できる形で先に用意する

### 8. 環境差の可視化方法

主な影響:

- 同じテンプレートでも環境が違うと結果差が説明しにくい
- 不具合報告時に再現条件を絞り込みにくい

取り得る方針:

- 方針1(Excel.Report.PDF、excelize)
  - Excel.Report.PDF や excelize は、コードコメント、テスト、最低限の診断で差異を扱う
  - 実装は比較的軽い
  - 環境差の比較には弱い
- 方針2(ReoGrid)
  - ReoGrid は、内部状態を豊富に持ち、印刷用レイアウトも構造として保持する
  - 後追い解析しやすい
  - 利用者向けの比較出力は別途必要になる
- 方針3(OysterReport 候補)
  - 診断に加えて `ReportDebugDumper` で環境情報付きダンプを出力する
  - 再現条件の比較と回帰検知に強い
  - ダンプ形式と運用ルールの設計が必要になる

現在の方針:

- `ReportDebugDumper` に `ReportMeasurementProfile` と実行環境情報を含める
- 診断とダンプの両方で差異を追えるようにする

## 3. 想定ユースケース

### 3.1 テンプレート帳票

- 請求書
- 納品書
- 見積書
- 作業報告書

### 3.2 アプリケーション編集

- `{{CustomerName}}` のようなプレースホルダ置換
- 明細行の値差し替え
- データ件数に応じたテンプレート行の追加
- 署名画像やロゴ画像の差し込み

レイアウトの微調整は原則としてプログラムからは行わず、Excel テンプレート側で行う。

## 4. 全体アーキテクチャ

```text
Excel(.xlsx)
  -> ExcelReader
  -> ReportWorkbook
     -> ReportSheet
        -> ReportRow / ReportColumn / ReportCell
        -> ReportPlaceholderText / ReportMergedRange / ReportImage
        -> ReportPageSetup / ReportHeaderFooter / ReportPrintArea / ReportPageBreak
  -> ReportWorkbook.AddSheet(...)
  -> ReportSheet.ReplacePlaceholder(...)
  -> ReportSheet.AddRows(...)
  -> PdfGenerator
     -> internal PdfRenderPlanner
     -> internal PdfRenderPlan
        -> internal PdfRenderSheetPlan
        -> internal PdfRenderPagePlan
        -> internal PdfCellRenderInfo / PdfBorderRenderInfo / PdfImageRenderInfo / PdfHeaderFooterRenderInfo
  -> ReportDebugDumper
  -> PDF
```

各責務は以下のように分離する。

- `ExcelReader`
  - ClosedXML から Excel を読む
  - Excel 単位を内部単位へ正規化する
  - 中間データ構造を組み立てる
- `ReportWorkbook` 系モデル
  - Excel 側の値、書式、印刷設定を主体に保持する
  - 外部からは読み取り専用で、変更は専用メソッドに限定する
- `ReportWorkbook.AddSheet`
  - シート追加を制御された変更として扱う
- `ReportSheet.ReplacePlaceholder`
  - Excel 上の特殊値を目印にセル表示文字列を差し替える
- `ReportSheet.AddRows`
  - データ件数に応じてテンプレート行を複製する
- `PdfGenerator`
  - 外部に公開する PDF 生成の入口とする
  - 内部では「レンダリング情報生成」と「PDF 描画」の 2 段階処理を行う
- `PdfRenderPlanner` と `PdfRenderPlan` 系モデル
  - `PdfGenerator` の内部でのみ使用する
  - 座標、テキスト測定、改ページ、罫線競合解決の結果を保持する
- `ReportDebugDumper`
  - 中間データ構造と PDF 生成前準備結果を JSON または Markdown でダンプ出力する
  - ただし internal なレンダリング情報クラスそのものは公開しない

### 4.1 名前空間設計

実装時の名前空間は `OysterReport` 配下で以下のように分割する。

- `OysterReport.Common`
  - 共通の基本型、列挙、結果型を配置する
- `OysterReport.Common.Geometry`
  - `ReportRect`、`ReportThickness`、`ReportLine` のような幾何型を配置する
- `OysterReport.Model`
  - `ReportWorkbook`、`ReportSheet`、`ReportCell` など Excel 主体の中間データ構造を配置する
- `OysterReport.Model.Printing`
  - `ReportPageSetup`、`ReportHeaderFooter`、`ReportPrintArea`、`ReportPageBreak` を配置する
- `OysterReport.Model.Images`
  - `ReportImage` や画像アンカー型を配置する
- `OysterReport.Reading`
  - `ExcelReader` と読み込みオプションを配置する
- `OysterReport.Reading.ClosedXml`
  - ClosedXML 依存の読み込み実装を隔離する
- `OysterReport.Writing.Pdf`
  - `PdfGenerator` と PDFsharp 依存の描画実装を配置する
- `OysterReport.Internal.Rendering`
  - `PdfRenderPlanner` と PDF 描画用レンダリング情報生成ロジックを配置する
- `OysterReport.Internal.Rendering.Plan`
  - `PdfRenderPlan`、`PdfRenderPagePlan` など internal な PDF 描画専用情報を配置する
- `OysterReport.Diagnostics`
  - `ReportDiagnostic`、`ReportDebugDumper` を配置する
- `OysterReport.Helpers`
  - 列幅換算、フォント解決補助、プレースホルダ解析などの補助機能を配置する
- `OysterReport.Internal.*`
  - 外部公開しない内部実装を配置する

方針:

- 中間データ構造は `OysterReport.Model` 側に閉じ込め、PDF 描画用の計算情報は `OysterReport.Internal.Rendering.Plan` へ分離する
- ClosedXML と PDFsharp への依存はそれぞれ `Reading.ClosedXml` と `Writing.Pdf` に閉じ込める
- 補助関数は `Helpers` へ集約するが、外部 API の主導権は `Model`、`Reading`、`Diagnostics`、`Writing.Pdf` に置く

## 5. 中間データ構造

### 5.1 基本方針

中間データ構造は、PDF 描画都合の計算結果ではなく、Excel 側の情報を主体に保持する。

- 保持するもの:
  - セル値、表示文字列、スタイル、行高、列幅、結合、印刷設定、ヘッダ、フッタ、画像アンカー
- 保持しないもの:
  - PDF 描画用に再計測した `TextBounds`
  - 共有辺競合解決後の罫線セグメント
  - ページごとの描画順序
  - 最終ページ分割結果

これらの PDF 描画専用情報は、`PdfGenerator` の内部で `PdfRenderPlanner` が生成する `PdfRenderPlan` 側に保持する。

### 5.2 ルートモデル

```csharp
public sealed class ReportWorkbook
{
    private readonly List<ReportSheet> sheets = [];
    private readonly List<ReportDiagnostic> diagnostics = [];

    public IReadOnlyList<ReportSheet> Sheets => sheets; // ワークブックに含まれるシート一覧
    public ReportMetadata Metadata { get; } // 帳票全体に関するメタデータ
    public ReportMeasurementProfile MeasurementProfile { get; } // 計測条件と環境差吸収設定
    public IReadOnlyList<ReportDiagnostic> Diagnostics => diagnostics; // 読み込み時点で収集した診断情報

    public ReportSheet AddSheet(string name);
    public void AddSheet(ReportSheet sheet);
}
```

```csharp
public sealed record ReportMetadata
{
    public string TemplateName { get; init; } = string.Empty; // テンプレート名
    public string? SourceFilePath { get; init; } // 読み込み元ファイルパス
    public DateTimeOffset? SourceLastWriteTime { get; init; } // 読み込み元の最終更新日時
    public string? Author { get; init; } // テンプレート作成者
}
```

```csharp
public sealed class ReportSheet
{
    private readonly List<ReportRow> rows = [];
    private readonly List<ReportColumn> columns = [];
    private readonly List<ReportCell> cells = [];
    private readonly List<ReportMergedRange> mergedRanges = [];
    private readonly List<ReportImage> images = [];
    private readonly List<ReportPageBreak> horizontalPageBreaks = [];
    private readonly List<ReportPageBreak> verticalPageBreaks = [];

    public string Name { get; } // シート名
    public ReportRange UsedRange { get; private set; } // 使用範囲
    public IReadOnlyList<ReportRow> Rows => rows; // 行定義一覧
    public IReadOnlyList<ReportColumn> Columns => columns; // 列定義一覧
    public IReadOnlyList<ReportCell> Cells => cells; // 使用範囲内のセル一覧
    public IReadOnlyList<ReportMergedRange> MergedRanges => mergedRanges; // 結合セル範囲一覧
    public IReadOnlyList<ReportImage> Images => images; // シート上の画像一覧
    public ReportPageSetup PageSetup { get; } // 印刷時のページ設定
    public ReportHeaderFooter HeaderFooter { get; } // ヘッダ、フッタ定義
    public ReportPrintArea? PrintArea { get; private set; } // 明示的な印刷範囲
    public IReadOnlyList<ReportPageBreak> HorizontalPageBreaks => horizontalPageBreaks; // 水平手動改ページ一覧
    public IReadOnlyList<ReportPageBreak> VerticalPageBreaks => verticalPageBreaks; // 垂直手動改ページ一覧
    public bool ShowGridLines { get; } // グリッド線表示フラグ

    public int ReplacePlaceholder(string markerName, string value);
    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values);
    public void AddRows(RowExpansionRequest request);
}
```

設計意図:

- `ReportWorkbook` / `ReportSheet` のプロパティは原則として外部から読み取り専用にする
- 変更は `AddSheet`、`ReplacePlaceholder`、`AddRows` など制御されたメソッド経由に限定する
- 帳票処理で必要な変更だけを許可し、行高、列幅、罫線、余白のようなレイアウト情報は Excel 側を正とする

### 5.3 セルモデル

```csharp
public sealed class ReportCell
{
    public int Row { get; private set; } // 1 始まりの行番号
    public int Column { get; private set; } // 1 始まりの列番号
    public string Address { get; private set; } = string.Empty; // A1 形式のセル番地
    public ReportCellValue Value { get; } // 元データ値
    public string SourceText { get; } = string.Empty; // Excel から読んだ元の表示文字列
    public string DisplayText { get; private set; } = string.Empty; // 現在の表示文字列
    public ReportPlaceholderText? Placeholder { get; } // プレースホルダ情報
    public ReportCellStyle Style { get; } = new(); // セルスタイル
    public ReportRect Bounds { get; private set; } // セル外枠の物理矩形
    public ReportMergeInfo? Merge { get; private set; } // 結合セル参加情報
}
```

```csharp
public sealed class ReportPlaceholderText
{
    public string MarkerText { get; } = string.Empty; // Excel 上の特殊値そのもの
    public string MarkerName { get; } = string.Empty; // アプリケーションから指定する識別子
    public string? ResolvedText { get; private set; } // 置換後の表示文字列
}
```

セルモデルの方針:

- `Value`、`SourceText`、`Style` は Excel ソース情報として保持する
- `DisplayText` はプレースホルダ置換結果を反映した現在値とする
- プレースホルダは「セルに設定された特殊値」を読み取り時に抽出し、`ReportSheet.ReplacePlaceholder` で置換する

### 5.4 スタイルモデル

```csharp
public sealed record ReportCellStyle
{
    public ReportFont Font { get; init; } = new(); // 文字フォント設定
    public ReportFill Fill { get; init; } = new(); // 背景塗りつぶし設定
    public ReportBorders Borders { get; init; } = new(); // 四辺の罫線設定
    public ReportAlignment Alignment { get; init; } = new(); // 水平、垂直配置設定
    public string? NumberFormat { get; init; } // Excel の表示書式
    public bool WrapText { get; init; } // 折り返し表示フラグ
    public double Rotation { get; init; } // 文字回転角度
    public bool ShrinkToFit { get; init; } // 縮小して全体表示するか
}
```

スタイルは Excel ソース情報として保持し、プログラムから直接変更する API は設けない。

### 5.5 行、列、画像、印刷ソース情報

```csharp
public sealed class ReportRow
{
    public int Index { get; private set; } // 1 始まりの行番号
    public double HeightPoint { get; private set; } // 行高(point)
    public double TopPoint { get; private set; } // シート先頭からの上端位置(point)
    public bool IsHidden { get; } // 非表示行かどうか
    public int OutlineLevel { get; } // アウトラインレベル
}
```

```csharp
public sealed class ReportColumn
{
    public int Index { get; private set; } // 1 始まりの列番号
    public double WidthPoint { get; private set; } // 列幅(point)
    public double LeftPoint { get; private set; } // シート左端からの左端位置(point)
    public bool IsHidden { get; } // 非表示列かどうか
    public int OutlineLevel { get; } // アウトラインレベル
    public double OriginalExcelWidth { get; } // Excel 上の元列幅値
}
```

```csharp
public sealed class ReportImage
{
    public string Name { get; } = string.Empty; // 画像識別名
    public ReportAnchorType AnchorType { get; } // アンカー種別
    public string FromCellAddress { get; } = string.Empty; // 開始セル番地
    public string? ToCellAddress { get; } // 終了セル番地
    public ReportOffset Offset { get; } // セル内オフセット
    public byte[]? ImageBytes { get; } // 元画像データ
}
```

```csharp
public sealed record ReportPageSetup
{
    public ReportPaperSize PaperSize { get; init; } // 用紙サイズ
    public ReportPageOrientation Orientation { get; init; } // 用紙向き
    public ReportThickness Margins { get; init; } // 本文余白
    public double HeaderMarginPoint { get; init; } // ヘッダ余白(point)
    public double FooterMarginPoint { get; init; } // フッタ余白(point)
    public int ScalePercent { get; init; } = 100; // 印刷倍率(%)
    public int? FitToPagesWide { get; init; } // 横方向の目標ページ数
    public int? FitToPagesTall { get; init; } // 縦方向の目標ページ数
    public bool CenterHorizontally { get; init; } // 水平中央寄せフラグ
    public bool CenterVertically { get; init; } // 垂直中央寄せフラグ
}
```

```csharp
public sealed record ReportHeaderFooter
{
    public bool AlignWithMargins { get; init; } // 余白に合わせるか
    public bool DifferentFirst { get; init; } // 先頭ページを別定義にするか
    public bool DifferentOddEven { get; init; } // 奇数偶数ページを別定義にするか
    public bool ScaleWithDocument { get; init; } = true; // 本文の拡大縮小に追従するか
    public string? OddHeader { get; init; } // 通常ページのヘッダ原文
    public string? OddFooter { get; init; } // 通常ページのフッタ原文
    public string? EvenHeader { get; init; } // 偶数ページのヘッダ原文
    public string? EvenFooter { get; init; } // 偶数ページのフッタ原文
    public string? FirstHeader { get; init; } // 先頭ページのヘッダ原文
    public string? FirstFooter { get; init; } // 先頭ページのフッタ原文
}
```

行、列、画像、印刷設定はセルの付属情報ではなく、Excel テンプレートから読んだ独立したソース情報として管理する。

### 5.6 幾何情報と計測プロファイル

PDF 出力と印刷再現の基準単位は point とする。

```csharp
public readonly record struct ReportRect
{
    public double X { get; init; } // 左上 X 座標(point)
    public double Y { get; init; } // 左上 Y 座標(point)
    public double Width { get; init; } // 幅(point)
    public double Height { get; init; } // 高さ(point)
}
```

```csharp
public sealed record ReportMeasurementProfile
{
    public double Dpi { get; init; } = 96; // 計測時に前提とする DPI
    public double MaxDigitWidth { get; init; } = 7.0; // 既定フォントでの最大数字幅
    public string DefaultFontName { get; init; } = string.Empty; // 既定フォント名
    public double DefaultFontSize { get; init; } = 11.0; // 既定フォントサイズ
    public double ColumnWidthAdjustment { get; init; } = 1.0; // 列幅換算の補正係数
}
```

これにより、環境差が出たときに以下を切り分けやすくする。

- Excel 列幅換算のずれ
- CJK フォントの文字幅差
- フォント未解決時のフォールバック差
- PDF 出力ホストごとの差異

### 5.7 内部 PDF レンダリング情報モデル (`internal`)

PDF 描画専用の情報は、中間データ構造とは別に internal な `PdfRenderPlan` として生成する。これらのクラスは `PdfGenerator` の内部実装専用であり、ライブラリ利用者へは直接公開しない。

```csharp
internal sealed record PdfRenderPlan
{
    public IReadOnlyList<PdfRenderSheetPlan> Sheets { get; init; } = Array.Empty<PdfRenderSheetPlan>(); // 解決済みシートレンダリング情報一覧
}
```

```csharp
internal sealed record PdfRenderSheetPlan
{
    public string SheetName { get; init; } = string.Empty; // 対象シート名
    public IReadOnlyList<PdfRenderPagePlan> Pages { get; init; } = Array.Empty<PdfRenderPagePlan>(); // ページ分割後のページ一覧
    public IReadOnlyList<PdfBorderRenderInfo> Borders { get; init; } = Array.Empty<PdfBorderRenderInfo>(); // 罫線競合解決後の罫線一覧
    public IReadOnlyList<PdfImageRenderInfo> Images { get; init; } = Array.Empty<PdfImageRenderInfo>(); // 画像の最終配置一覧
}
```

```csharp
internal sealed record PdfRenderPagePlan
{
    public int PageNumber { get; init; } // 1 始まりのページ番号
    public ReportRect PageBounds { get; init; } // 用紙全体の矩形
    public ReportRect PrintableBounds { get; init; } // 余白を除いた印字可能領域
    public PdfHeaderFooterRenderInfo HeaderFooter { get; init; } = new(); // 当該ページのヘッダ、フッタ描画情報
    public IReadOnlyList<PdfCellRenderInfo> Cells { get; init; } = Array.Empty<PdfCellRenderInfo>(); // 当該ページに描画するセル一覧
}
```

```csharp
internal sealed record PdfCellRenderInfo
{
    public string CellAddress { get; init; } = string.Empty; // 対象セルの番地
    public ReportRect OuterBounds { get; init; } // セル外枠の最終矩形
    public ReportRect ContentBounds { get; init; } // 内容描画領域の最終矩形
    public ReportRect TextBounds { get; init; } // テキスト描画に使う最終矩形
    public bool IsMergedOwner { get; init; } // 結合セルの代表セルかどうか
    public bool IsClipped { get; init; } // 描画時にクリップが必要かどうか
}
```

```csharp
internal sealed record PdfBorderRenderInfo
{
    public ReportLine Line { get; init; } // 描画する線分
    public ReportBorderStyle Style { get; init; } // 線分に適用する罫線スタイル
    public string OwnerCellAddress { get; init; } = string.Empty; // この線分の由来となる代表セル番地
}
```

```csharp
internal sealed record PdfHeaderFooterRenderInfo
{
    public string? HeaderText { get; init; } // 当該ページに描画するヘッダ文字列
    public string? FooterText { get; init; } // 当該ページに描画するフッタ文字列
    public ReportRect HeaderBounds { get; init; } // ヘッダ描画領域
    public ReportRect FooterBounds { get; init; } // フッタ描画領域
}
```

この分離により、以下が可能になる。

- 中間データ構造を Excel ソース情報として保ちやすい
- プレースホルダ置換や行追加の後に、レンダリング情報だけを再生成できる
- ページ分割、罫線競合、テキスト計測の責務を PDF 描画フェーズ側へ寄せられる
- デバッグ時に「ソース情報」と「描画計算結果」を分けて比較できる

### 5.8 可変性ポリシー

中間データ構造の可変性は以下のように制限する。

- `ReportWorkbook`、`ReportSheet`、`ReportCell` の公開プロパティは原則読み取り専用とする
- シート追加、プレースホルダ置換、行追加はメソッド経由でのみ行う
- 行高、列幅、結合、罫線、余白などのレイアウト情報はプログラムから直接編集しない
- PDF 描画専用の計算結果は internal な `PdfRenderPlan` 側へ保持し、中間データ構造へ戻さない

この方針により、テンプレートで作り込んだ Excel レイアウトを保護しつつ、帳票生成に必要な最小限の編集だけを許可できる。

## 6. Excel 読み込み仕様

### 6.1 入力

`ExcelReader` は以下を受け付ける。

- ファイルパス
- `Stream`

API 例:

```csharp
public sealed class ExcelReader
{
    public ReportWorkbook Read(string filePath, ExcelReadOptions? options = null);
    public ReportWorkbook Read(Stream stream, ExcelReadOptions? options = null);
}
```

### 6.2 読み取り対象

- ワークブック内の対象シート
- 使用範囲
- 行高、列幅
- セル値
- 表示文字列
- 特殊値プレースホルダ
- セルスタイル
- 結合セル
- 画像
- 非表示行、非表示列
- アウトラインレベル
- 印刷設定
- ヘッダ、フッタ
- 印刷範囲
- 手動改ページ

### 6.3 数式セル

数式の再計算は行わない。v1 では以下の優先順位で表示値を決定する。

1. Excel ファイルに保存されている計算済み表示値
2. 取得できるセル値
3. 取得不能時は警告を記録し、空文字または数式文字列を採用

### 6.4 行高・列幅の変換

Excel の行高・列幅はそのままでは PDF 描画に使いにくいため、読み込み時に point に正規化する。

- 行高は Excel の point 値をそのまま利用
- 列幅は Excel 列幅から point へ変換する
- 変換結果は列ごとにキャッシュする

列幅の再現は誤差が出やすいため、以下を仕様とする。

- 変換式は OysterReport 内で一元管理する
- 実測差が大きい場合に備えて調整係数をオプションで持てるようにする
- 将来的にフォント依存の補正を追加できるようにする

### 6.5 画像

セルにアンカーされた画像は以下を保持する。

- 元画像データまたは取得可能な参照
- 開始セル、終了セル、オフセット
- 配置モード
- 元のアンカーセル情報

v1 では画像の回転や高度なトリミングは限定対応とする。

### 6.6 非表示行、非表示列

非表示行、非表示列の扱いは以下とする。

- 読み込み時には hidden フラグをそのまま保持する
- 中間データ構造上では `ReportRow.IsHidden`、`ReportColumn.IsHidden` として参照できる
- v1 の PDF レンダリングでは本文描画対象から除外する
- 改ページ、印刷可能領域、画像アンカー解決も、レンダリング時には除外後の可視行、可視列を前提に計算する

参照ライブラリ比較:

- ReoGrid は `IsCellVisible` や hidden row / column の概念を持ち、印刷系ロジックから可視性を考慮できる
- excelize は `hiddenRows` / `hiddenColumns` や行列 hidden 属性を保持できる
- Excel.Report.PDF の手元確認範囲では、非表示行、非表示列を切り替える公開方針は強く出ていない

### 6.7 ヘッダ、フッタ

ヘッダ、フッタは Excel ソース情報として `ReportHeaderFooter` に保持する。

- 奇数ページ、偶数ページ、先頭ページの定義を保持する
- 余白連動、拡大縮小連動のフラグを保持する
- ページ番号のような実ページ依存値は、ソースではトークンのまま保持し、レンダリング時に解決する

参照ライブラリ比較:

- Excel.Report.PDF は余白計算でヘッダ、フッタ余白を取り込んでいる
- excelize は `SetHeaderFooter` と `AddHeaderFooterImage` を持ち、Excel 構造としては広く保持できる
- OysterReport v1 は本文との競合制御を優先し、まずはテキストヘッダ、フッタを対象とする

## 7. 編集仕様

### 7.1 基本方針

中間データ構造は、レイアウトをプログラムから直接編集する前提にはしない。Excel 操作 API をそのまま模倣するのではなく、帳票処理で必要な以下の操作に限定する。

- シート追加
- プレースホルダ解決
- データ件数に応じた行追加

行高、列幅、結合、罫線、余白などのレイアウト情報は原則として Excel テンプレート側で編集し、プログラムからは変更しない。

編集後は `PdfGenerator` 内部のレンダリング計画生成が再実行され、既存の internal レンダリング情報は破棄して再構築する。

### 7.2 必須操作

- シート追加
- プレースホルダ解決
- データ件数に応じたテンプレート行の追加
- 画像差し替え

API 例:

```csharp
public sealed class ReportWorkbook
{
    public ReportSheet AddSheet(string name);
    public void AddSheet(ReportSheet sheet);
}
```

```csharp
public sealed class ReportSheet
{
    public int ReplacePlaceholder(string markerName, string value);
    public int ReplacePlaceholders(IReadOnlyDictionary<string, string?> values);
    public void AddRows(RowExpansionRequest request);
}
```

```csharp
public sealed record RowExpansionRequest
{
    public string SheetName { get; init; } = string.Empty; // 行追加対象のシート名
    public ReportRange TemplateRowRange { get; init; } // 複製元となるテンプレート行範囲
    public IReadOnlyList<IReadOnlyDictionary<string, object?>> Items { get; init; } = Array.Empty<IReadOnlyDictionary<string, object?>>(); // 追加行ごとの差し込みデータ
}
```

`ReportWorkbook` / `ReportSheet` のメソッドは公開 API として提供し、内部的には helper や internal service を使って整合性を維持してよい。

### 7.3 プレースホルダ

v1 の標準プレースホルダ形式は、セル全体に設定する特殊値 `{{Name}}` とする。

理由:

- Excel 上で視認しやすい
- `ReportSheet.ReplacePlaceholder("Name", "...")` のようにメソッドから指定しやすい
- セル内埋め込み置換よりも Excel 主体のソースモデルを単純に保ちやすい

置換ルール:

- v1 はセル全体一致の特殊値を対象とする
- 読み込み時に `{{CustomerName}}` は `ReportPlaceholderText` として抽出する
- 未解決プレースホルダは既定ではそのまま残す
- オプションで未解決時に例外化できるようにする
- 置換結果は `ReportCell.DisplayText` と `ReportPlaceholderText.ResolvedText` に保持する

### 7.4 行追加

データ件数に応じて行を追加する操作は想定する。ただし、既存の行や列のレイアウト情報を直接変更する API は提供しない。

方針:

- Excel 側で「繰り返し元となるテンプレート行」を定義する
- `ReportSheet.AddRows` がそのテンプレート行を複製し、同一シート内の source model を更新する
- 行追加に伴う行位置、セル番地、結合セル、画像アンカー、使用範囲は再計算する
- PDF 描画用のページ分割や `TextBounds` は更新せず、後段で `PdfGenerator` 内部が再生成する

この方式により、レイアウトの編集責任は Excel に残しつつ、データ件数に応じた帳票伸長を行える。

### 7.5 明細行の繰り返し

v1 では Excel テンプレート内の独自 DSL を必須にしない。

代わりに以下の方針とする。

- まずは Excel テンプレート側で繰り返し元行を指定し、`ReportSheet.AddRows` で複製できればよい
- ライブラリの標準機能としては、将来 `TemplateProcessor` を追加できる余地を残す

この判断により、Excel.Report.PDF のような記号ベース帳票機能を将来拡張で取り込める一方、コアモデルは単純に保てる。

## 8. PDF 生成仕様

### 8.1 API

```csharp
public sealed class PdfGenerator
{
    public void Generate(
        ReportWorkbook workbook,
        Stream output,
        PdfGenerateOptions? options = null);

    internal PdfRenderPlan BuildRenderPlan(
        ReportWorkbook workbook,
        PdfGenerateOptions options);

    internal void WritePdf(
        ReportWorkbook workbook,
        PdfRenderPlan renderPlan,
        Stream output,
        PdfGenerateOptions options);
}
```

2 フェーズ設計:

1. `PdfGenerator.Generate` の内部で `BuildRenderPlan` が `PdfRenderPlanner` を使って internal な `PdfRenderPlan` を生成する
2. 続けて `WritePdf` が `PdfRenderPlan` と `ReportWorkbook` を使って PDF を描画する

利用者は `PdfGenerator.Generate` だけを呼び出せばよく、2 段階処理はすべて `PdfGenerator` の内部責務とする。

### 8.2 描画対象

- セル背景
- 罫線
- テキスト
- 画像
- ヘッダ
- フッタ
- ページ番号

非表示行、非表示列は v1 では描画対象から除外する。

描画順は以下とする。

1. ページ背景
2. セル背景
3. 罫線
4. テキスト
5. 画像
6. ヘッダ、フッタ、ページ番号

### 8.3 テキスト描画

以下を考慮する。

- 水平配置
- 垂直配置
- 折り返し
- 縮小して全体を表示
- 回転
- 結合セル内配置

文字描画は `ReportCell.DisplayText` から得られる最終文字列を使う。v1 では Excel のオートフィット再計算はしない。

### 8.4 罫線

罫線は各辺単位で保持し、PDF でも各辺単位で描画する。

- 線種
- 線幅
- 色

二重線や破線は、PDFsharp で再現困難な場合に近似描画を許容する。

### 8.5 改ページ

改ページは以下の優先順で決定する。

1. Excel の明示的な印刷範囲
2. Excel の明示的な手動改ページ
3. 用紙サイズ、余白、倍率に基づく自動改ページ

自動改ページでは以下を考慮する。

- 用紙サイズ
- 向き
- 印字可能領域
- ヘッダ余白、フッタ余白
- 行高、列幅の累積
- 結合セルを途中で分断しないこと
- 非表示行、非表示列を除外した可視領域

必要に応じて横方向、縦方向の両方でページ分割する。

### 8.6 ページ設定

以下を `ReportPageSetup` に保持する。

- 用紙サイズ
- 向き
- 左右上下余白
- ヘッダ余白、フッタ余白
- 印刷倍率
- 横何ページ、縦何ページに収めるか
- 中央寄せ

### 8.7 ページ番号

v1 ではページ番号トークンをセルテキストそのものに埋め込むのではなく、ヘッダ / フッタ描画機能として提供する。

将来拡張として以下のような組み込みトークンを許可できる設計にする。

- `{PageNumber}`
- `{TotalPages}`

### 8.8 ヘッダ、フッタ

ヘッダ、フッタは `ReportHeaderFooter` を元に、`PdfGenerator` 内部のレンダリング計画生成フェーズがページ単位の `PdfHeaderFooterRenderInfo` を生成する。

- 先頭ページ、奇数ページ、偶数ページの切り替えをレンダリング時に判断する
- 本文の `PrintableBounds` と競合しないよう、ヘッダ、フッタ余白を含めてページ領域を解決する
- ページ番号トークンのような実ページ依存値は、レンダリングフェーズで最終文字列へ変換する
- ヘッダ、フッタ画像は v1 では限定対応とし、必要時は診断へ記録する

## 9. フォント仕様

### 9.1 基本方針

PDFsharp の制約上、使用フォントは `IFontResolver` 相当の仕組みで解決する。

OysterReport は独自の抽象化を持つ。

```csharp
public interface IReportFontResolver
{
    ReportFontResolveResult Resolve(ReportFontRequest request);
}
```

### 9.2 フォント解決ルール

- Excel で指定されたフォント名を第一候補とする
- 太字、斜体、文字集合を考慮する
- 該当フォントが解決できない場合はフォールバックを適用する
- どのフォントへフォールバックしたかを診断情報に残す

### 9.3 診断

フォント周りは見た目差異の主因になるため、以下を診断可能にする。

- 未解決フォント一覧
- フォールバック発生回数
- フォールバック先フォント
- 文字描画不能の可能性がある文字種

## 10. 診断とエラー処理

### 10.1 原則

- 読み込み不能な致命エラーは例外
- 軽微な再現差異は警告

### 10.2 警告例

- 未解決フォント
- 未対応の罫線種別
- 未対応の画像形式
- 数式結果未保存
- 条件付き書式を無視

### 10.3 診断オブジェクト

```csharp
public sealed class ReportDiagnostic
{
    public ReportDiagnosticSeverity Severity { get; init; } // 重大度
    public string Code { get; init; } // 診断コード
    public string Message { get; init; } // 利用者向け診断メッセージ
    public string? SheetName { get; init; } // 関連シート名
    public string? CellAddress { get; init; } // 関連セル番地
}
```

`ExcelReader` と `PdfGenerator` は診断情報を返却または参照可能にする。

### 10.4 デバッグダンプ

レイアウト不一致や変換不具合を調査しやすくするため、中間データ構造と PDF 生成前準備結果のダンプ出力機能を持つ。

```csharp
public sealed class ReportDebugDumper
{
    public void DumpWorkbook(
        ReportWorkbook workbook,
        Stream output,
        ReportDumpFormat format = ReportDumpFormat.Json);

    public void DumpPdfPreparation(
        ReportWorkbook workbook,
        Stream output,
        PdfGenerateOptions? options = null,
        ReportDumpFormat format = ReportDumpFormat.Json);
}
```

ダンプ対象:

- シート、行、列、セルの基本情報
- プレースホルダ解決状態
- 結合セル範囲
- 画像アンカー情報
- ヘッダ、フッタ定義
- `PdfGenerator` 内部で生成される改ページ結果
- セル描画矩形、罫線競合解決結果、画像最終配置、ヘッダ、フッタ配置の要約
- `ReportMeasurementProfile`
- 実行環境情報

実行環境情報の例:

- OS
- プロセスアーキテクチャ
- 現在カルチャ
- 使用フォントリゾルバ
- 主要フォールバック結果

主な用途:

- レイアウト崩れの原因調査
- テスト失敗時の比較
- Excel 読み込み結果と PDF 生成前レンダリング情報の差分確認

## 11. 推奨公開 API

```csharp
public sealed class ExcelReadOptions
{
    public string[]? TargetSheets { get; set; } // 読み込み対象シート名一覧
    public bool IncludeImages { get; set; } = true; // 画像も読み込むか
}
```

```csharp
public sealed class PdfGenerateOptions
{
    public IReportFontResolver? FontResolver { get; set; } // PDF 描画時に使うフォントリゾルバ
    public bool StrictMode { get; set; } // 未解決要素をエラー扱いする厳格モードか
    public double MinimumReadableScale { get; set; } = 0.5; // 許容する最小拡大縮小率
    public bool EmbedDocumentMetadata { get; set; } = true; // PDF 文書メタデータを書き込むか
    public bool CompressContentStreams { get; set; } = true; // PDF コンテンツを圧縮するか
}
```

```csharp
public sealed class OysterReportEngine
{
    public ReportWorkbook Read(string filePath, ExcelReadOptions? options = null);
    public void GeneratePdf(ReportWorkbook workbook, Stream output, PdfGenerateOptions? options = null);
    public void DumpWorkbook(ReportWorkbook workbook, Stream output, ReportDumpFormat format = ReportDumpFormat.Json);
    public void DumpPdfPreparation(ReportWorkbook workbook, Stream output, PdfGenerateOptions? options = null, ReportDumpFormat format = ReportDumpFormat.Json);
}
```

利用者にとっては `OysterReportEngine` を入口にしつつ、低レベル API として `ExcelReader` / `ReportDebugDumper` / `PdfGenerator` を個別利用できる構成が望ましい。プレースホルダ置換や行追加は `ReportWorkbook` / `ReportSheet` のメソッドで行い、PDF 準備用の internal クラスは公開 API に含めない。

## 12. 非機能要件

### 12.1 再現性

- 同じ入力、同じフォント解決条件なら同じ PDF を生成できること
- 罫線、余白、改ページが実行ごとに変動しないこと

### 12.2 性能

- 複数ページ帳票を想定し、不要なセル再計算を避ける
- PDF 描画時にページ単位で処理できる設計にする
- 行列の累積位置は事前計算しておく

### 12.3 保守性

- Excel 読み込み、モデル編集、PDF 描画を分離する
- ClosedXML / PDFsharp への依存を境界に閉じ込める
- 将来の別描画エンジン差し替え余地を残す

## 13. テスト方針

最低限、以下を自動テスト対象にする。

- 行高、列幅の変換
- 結合セル矩形の算出
- 罫線描画位置
- 余白を含むページ矩形計算
- 手動改ページと自動改ページの優先順位
- プレースホルダ置換
- 行追加後の行位置、結合セル、画像アンカー再計算
- フォントフォールバック診断
- ダンプ出力のスナップショット比較

目視確認用として、以下のゴールデンファイルを用意する。

- 単票
- 複数ページ帳票
- 結合セルを多用した帳票
- 日本語フォント帳票
- 画像差し込み帳票

## 14. 実装順序

1. `ReportWorkbook` / `ReportSheet` の読み取り専用プロパティと制御メソッドの境界を確定する
2. `ExcelReader` で使用範囲、行高、列幅、セル値、結合セルまで読み込む
3. `ReportWorkbook.AddSheet`、`ReportSheet.ReplacePlaceholder`、`ReportSheet.AddRows` を実装する
4. `PdfGenerator` 内部のレンダリング計画生成で座標、罫線、改ページ、ヘッダ、フッタの解決を実装する
5. `PdfGenerator` で背景、罫線、文字、画像、ヘッダ、フッタの描画を実装する
6. フォントリゾルバ、診断機構、`ReportDebugDumper` を実装する

## 15. 仕様決定メモ

### 15.1 Excel.Report.PDF から取り入れる考え方

- Excel の印刷設定を PDF 側のページ設定へ引き継ぐ
- フォント解決を外部注入可能にする
- テンプレート帳票に必要な差し込み機能を持たせる
- 複数ページ帳票を最初から考慮する

### 15.2 ReoGrid から取り入れる考え方

- セル範囲の物理座標を明示的に持つ
- 印刷範囲と改ページを独立したモデルとして管理する
- 自動改ページは行高・列幅の累積で決定する

### 15.3 Excelize から取り入れる考え方

- 行、列、画像の位置計算を明示的に扱う
- 結合セルは矩形として正規化する
- 列幅変換やテキスト幅計算を専用関数へ閉じ込める

### 15.4 DioDocs for Excel を踏まえた方針

- 利用者は「Excel テンプレートから帳票 PDF を安定して出したい」という期待を持つため、OysterReport も印刷設定と見た目再現を最優先にする
- ただし v1 は商用製品と同等の完全互換ではなく、帳票用途に必要な機能へ対象を絞る

## 16. 参考ライブラリ比較と強化ポイント

### 16.1 Excel.Report.PDF から見える強みと不足

強み:

- Excel の印刷設定、余白、中央寄せ、倍率を PDF 出力へ直接反映している
- セル背景、罫線、文字、画像の描画順が整理されている
- 共有辺の罫線競合を優先順位で解決している
- `#FitColumn` や `#PageCount` のように帳票実用上の機能を先に用意している

不足:

- 構造化された中間モデルが薄く、読み込み結果をアプリケーション都合で再構成しにくい
- セル内容の再現が「描画時の即時計算」に寄っており、描画前に検証しにくい
- 列幅換算やスケーリングが固定ロジック寄りで、環境差の吸収余地が小さい
- `#Empty` のような特殊指示がセル文字列へ埋め込まれており、構造化しづらい

OysterReport の採用方針:

- 共有辺の罫線競合解決は Excel.Report.PDF に近い方針を採用し、辺ごとの優先順位判定を行う
- 背景、罫線、文字、画像の描画順は Excel.Report.PDF と同様にレイヤー分離する
- `#FitColumn` が担っている「印字可能領域に対する横方向スケール決定」は、独自記号ではなく `PdfGenerator` 内部の標準ロジックとして取り込む
- ページ番号やヘッダ、フッタ競合はセル文字列ではなくレンダリング情報で解く

検討が必要な項目:

- Excel.Report.PDF は描画時に多くの判断を行っているが、OysterReport は描画前にレンダリング情報を作るため、責務分離の実装コストが増える
- ページ番号差し込みのような描画後確定情報は、セル文字列へ埋め込む方式より、ヘッダ、フッタ機構で扱う方が構造的には健全だが、Excel テンプレート互換の見え方は少し下がる

### 16.2 ReoGrid から見える強みと不足

強み:

- セル矩形、範囲矩形、テキスト矩形を別々に保持している
- `PrintableRange`、手動改ページ、自動改ページを分離して扱っている
- 結合セルの開始、終了、境界更新を内部整合込みで管理している
- 印刷前の計算状態とページ分割結果を分けている

不足:

- Excel テンプレート読み込みと PDF 帳票生成を一体で提供する設計ではない
- ヘッダ、フッタや Excel テンプレート固有の差し込みモデルは主対象ではない
- そのまま導入すると帳票用途に対して API が広すぎる

OysterReport の採用方針:

- `ReportWorkbook` と internal な `PdfRenderPlan` の二層化は ReoGrid の印刷前計算分離に近づける
- `TextBounds` と印刷用テキスト配置は ReoGrid と同様に独立したレンダリング情報として保持する
- `PrintableRange`、手動改ページ、自動改ページの分離も ReoGrid 寄りに採用する
- 結合セルの整合更新は、開始セル、終了セル、範囲矩形を常に同期させる ReoGrid に近い方針とする

検討が必要な項目:

- OysterReport では ReoGrid の UI や汎用編集機構は比較対象に含めず、印刷範囲、改ページ、テキスト配置、結合整合の考え方だけを抽出する
- 帳票用途で必要な PDF / 印刷関連部分のみに限定して取り込む

### 16.3 Excelize から見える強みと不足

強み:

- 列幅換算、画像アンカー位置、結合セル矩形を専用関数へ分離している
- 結合セルの重なりを正規化し、範囲として扱っている
- オブジェクト配置をセル基準とオフセット基準で計算している
- 近似計算である点や非対応条件をコメントで明示している

不足:

- AutoFit は近似であり、結合セルを考慮しない前提がある
- Go ライブラリなので、そのまま .NET の測定基盤には流用できない
- PDF 再現よりも Excel ファイル編集全般が主目的である

OysterReport の採用方針:

- 列幅換算は Excelize のように専用関数へ閉じ込め、計測仮定を明示する
- 画像は Excelize に近く、ソース側で `Anchor + Offset` を保持し、レンダリング側で最終矩形へ解決する
- 結合セルは Excelize と同様に矩形として正規化し、重なりや破損を診断できるようにする

検討が必要な項目:

- Excelize の列幅換算や AutoFit は近似式ベースであり、.NET の実測結果と一致しない可能性がある
- そのため OysterReport は Excelize の「専用計算関数に閉じ込める」設計は採用するが、計算式自体は差し替え可能にする

### 16.4 DioDocs for Excel から見える強みと不足

強み:

- Excel テンプレートに対してデータバインドや帳票生成を行う方向性が明確
- Excel から PDF へ、ページ設定や図形、チャートを含めて保存する製品思想を持つ

不足:

- 商用製品であり、内部設計をそのまま参照できない
- 完全互換志向の設計は v1 の実装コストを大きく押し上げる

OysterReport の採用方針:

- 製品思想としては DioDocs for Excel に近く、「Excel テンプレートを帳票としてできるだけそのまま PDF 化する」方向を採る
- ページ設定、フォント、図形、画像、改ページをレイアウト再現の中心テーマとして扱う
- 将来拡張を見据えて、テンプレート処理や図形を保持できる中間構造を用意する

検討が必要な項目:

- DioDocs for Excel のような高忠実度再現は、図形、チャート、条件付き書式、印刷エンジン全体の再現まで含みやすく、v1 の実装量を大幅に増やす
- そのため v1 は帳票で重要度の高い行列、文字、罫線、画像、改ページへ集中し、グラフや高度な図形は将来対応とする

### 16.5 総合的に強化すべき点

レイアウト忠実度を優先した場合、OysterReport で採用すべき方針は以下。

- 改ページ、印刷範囲、印刷倍率、中央寄せの扱いは ReoGrid と Excel.Report.PDF に近づける
- 罫線競合解決と描画順は Excel.Report.PDF に近づける
- テキスト矩形と印刷矩形の分離は ReoGrid に近づける
- 列幅換算、画像アンカー、結合セル矩形正規化は Excelize に近づける
- 製品としての目標は DioDocs for Excel に近い「テンプレート忠実再現」に置く

そのうえで、OysterReport 独自に必要な補強は以下。

- 読み取り専用プロパティ中心の中間モデルと描画用レンダリング情報の二層化
- 列幅、フォント、テキスト計測の仮定を `ReportMeasurementProfile` と診断へ露出すること
- `ContentBounds`、`TextBounds`、ヘッダ、フッタ描画領域をレンダリング情報として保持できること
- 手動改ページ、印刷範囲、自動改ページの決定理由を保持すること

## 17. レイアウト再現仕様

### 17.1 基本方針

レイアウト再現は「セルを順に描けばよい」という問題ではない。OysterReport では、Excel の内容をそのまま描画するのではなく、以下の 4 段階で再現する。

1. Excel からソース情報を抽出する
2. point 基準の中間データ構造へ正規化する
3. PDF 用のレンダリング情報を生成する
4. レンダリング情報と中間データ構造から PDF を描画する

この方式により、Excel 読み込み、アプリケーション編集、最終描画を分離する。

忠実な再現を優先するため、以下の原則を採る。

- 罫線競合、描画順、縮尺決定は Excel.Report.PDF 寄りにする
- テキスト矩形、印刷矩形、改ページ管理は ReoGrid 寄りにする
- 列幅換算、画像アンカー、結合セル矩形正規化は Excelize 寄りにする
- 目標品質は DioDocs for Excel が想定している「Excel テンプレートをそのまま帳票化する体験」に近づける

### 17.2 再現パイプライン

```text
Excel source
  -> 値 / 表示文字列 / スタイル / 行列寸法 / 印刷設定 / ヘッダ / フッタ / 画像アンカー抽出
  -> ReportWorkbook
  -> ReportWorkbook.AddSheet / ReportSheet.ReplacePlaceholder / ReportSheet.AddRows
  -> PdfGenerator(internal planning)
     -> 行列累積位置の確定
     -> 結合セル矩形の確定
     -> テキスト計測
     -> 罫線競合解決
     -> 印刷範囲と改ページの確定
     -> ヘッダ / フッタ文字列のページ別解決
     -> ページごとの描画項目生成
  -> internal PdfRenderPlan
  -> PdfGenerator
```

### 17.3 再現対象

レイアウト再現対象は以下。

- 行高、列幅、セル矩形
- 余白、用紙サイズ、向き、倍率、中央寄せ
- 結合セルの占有範囲
- 背景色、文字色、罫線
- テキスト配置、折り返し、回転、縦書き、縮小表示
- 画像のアンカー位置
- 印刷範囲、手動改ページ、自動改ページ
- ヘッダ、フッタ、ページ番号
- 非表示行、非表示列の除外

### 17.4 単位と座標の扱い

再現はすべて point 基準で行う。

- 行高は point として取り込む
- 列幅は Excel 列幅から point へ変換する
- 行、列の累積位置をシート単位で保持する
- 画像オフセットも point に正規化する

列幅変換は最も誤差が出やすい箇所なので、以下を同時に保持する。

- 元の Excel 列幅
- 変換後の point 値
- 使用した `ReportMeasurementProfile`

これにより、環境差の切り分けが可能になる。

OysterReport の方針:

- 列幅換算は Excelize のように専用関数へ隔離する
- ただし換算式は固定し切らず、`ReportMeasurementProfile` で微調整可能にする

検討が必要な項目:

- Excel 実機に近づけるほど、既定フォントや文字幅の実測に依存する
- その結果、フォント未解決時や Linux / Windows 差異で列幅が変動しやすくなる
- このため v1 では「計算式を固定して隠す」よりも、「仮定を露出して補正可能にする」方を優先する

### 17.5 テキスト再現

テキストは以下の 3 つの矩形を分けて扱う。

- `OuterBounds`
  - セルの外枠
- `ContentBounds`
  - パディングや結合セルを考慮した内容領域
- `TextBounds`
  - 実測後の描画矩形

再現手順:

1. `ReportCell.DisplayText` から描画対象文字列を決定する
2. フォント解決を行う
3. 文字列計測を行う
4. 配置、折り返し、回転、縦書き、縮小表示を反映して `TextBounds` を確定する

今の想定仕様で問題になりうる点:

- CJK フォントの実測幅が環境によってずれる
- `ShrinkToFit` と折り返しが同時指定されたときの優先順位
- 結合セル内での中央寄せと複数行表示
- セル幅を超えるテキストの隣接セルオーバーフロー

他ライブラリの扱い:

- ReoGrid は `TextBounds` と `PrintTextBounds` を別管理して印刷用レイアウトを持つ
- Excelize は AutoFit を近似計算として扱い、結合セル未対応であることを明示している
- Excel.Report.PDF は描画時に回転、縦書き、ページ番号差し込みを処理している

OysterReport の方針:

- ReoGrid に近く `TextBounds` を明示的に保持する
- Excel.Report.PDF に近く回転、縦書き、縮小表示は描画要件として初期から扱う
- 実測結果とクリップ発生を診断に残す
- 隣接セルオーバーフローは、忠実再現のためには Excel 互換の隣接セル侵食ロジックが必要だが、v1 では限定対応とし、診断で可視化する

検討が必要な項目:

- Excel の隣接セルオーバーフローは、隣接セルの空状態、結合状態、罫線、回転文字と相互作用するため、実装が複雑
- 完全対応には描画前に横方向占有判定を再計算する必要があり、改ページや画像と衝突しやすい
- そのため v1 では `TextBounds` とクリップ情報を保持しつつ、オーバーフロー再現は段階導入とする

### 17.6 罫線再現

罫線はセルに属しているように見えるが、実際には共有辺の競合解決が必要である。

問題になりうる点:

- 隣接セルが異なる罫線種別を持つ
- 同じ辺を二重描画すると濃く見える
- 二重線、破線、Hairline の近似表現が必要

他ライブラリの扱い:

- Excel.Report.PDF は罫線種別に優先順位を与え、共有辺の描画主体を決めている
- ReoGrid は境界線をセルデータとは別に管理し、結合時に境界を切り直している

OysterReport の方針:

- Excel.Report.PDF に近く、共有辺ごとに優先順位を解決してから一度だけ描画する
- ReoGrid に近く、結合セル更新時に境界セグメントの整合を取り直す
- 描画直前に `PdfBorderRenderInfo` へ正規化する
- PDF で完全再現不能な線種は近似描画し、必要なら診断へ記録する

検討が必要な項目:

- Excel の斜め罫線、複雑な二重線、Hairline は PDFsharp での完全再現が難しい
- 線の優先順位を忠実に再現するほど、共有辺判定と結合セル境界の調整が複雑になる
- このため v1 は外周罫線、共有辺罫線、二重線近似を優先し、特殊線種は段階導入とする

### 17.7 結合セル再現

結合セルは「複数セルの見た目」ではなく「単一の描画矩形」として扱う。

問題になりうる点:

- 結合セルの途中で改ページされる
- 結合セル上でテキスト配置や画像配置が崩れる
- 結合セル内の内側罫線を誤って描画する

他ライブラリの扱い:

- ReoGrid は開始セル、終了セル、範囲境界を保持し、結合時に範囲全体の整合を更新する
- Excelize は結合セルを矩形として正規化し、重なりを統合する
- Excel.Report.PDF は結合セルの先頭セルに統合サイズを持たせて描画している

OysterReport の方針:

- ReoGrid と Excelize に近く、結合セルは開始点、終了点、矩形を常に同期させる
- `ReportMergedRange` と `PdfCellRenderInfo.IsMergedOwner` を持つ
- 改ページ時は結合セルを途中分断しないのを原則とする
- 分断回避不能時は `LayoutSplitMergedRange` 診断を出す

検討が必要な項目:

- ページ境界をまたぐ巨大な結合セルは、どの帳票でも再現が難しい
- 完全に分断禁止にすると極端な縮小や空白ページが発生する場合がある
- そのため v1 では「分断しない」を第一優先にしつつ、回避不能時のみ警告付きで近似処理を認める

### 17.8 画像、図形の再現

v1 での主対象は画像とする。

問題になりうる点:

- セルサイズ変更後にアンカー再計算が必要
- 結合セル上にある画像の位置ずれ
- OneCellAnchor / TwoCellAnchor 的な差異

他ライブラリの扱い:

- Excelize はセル幅、高さ、オフセットから画像の終端セルとオフセットを計算している
- DioDocs for Excel は図形やチャートを含む PDF 保存を前提にしている

OysterReport の方針:

- Excelize に近く、画像はアンカーセルとオフセットからレンダリング時に最終矩形を算出する
- `ReportImage` にアンカー種別、開始セル、終了セル、オフセットを持たせ、internal な `PdfImageRenderInfo` に最終矩形を持たせる
- DioDocs for Excel の方向性を見据え、将来的に図形やチャートも載せられるレイヤー構造にする
- v1 では画像を優先し、図形は将来拡張とする

検討が必要な項目:

- 図形やチャートまで含めて高忠実度再現を目指すと、Open XML 図形モデルの実装が大きくなる
- 帳票用途ではまず画像の再現性が重要度高いため、v1 は画像に集中する

### 17.9 改ページと印刷設定の再現

改ページはレイアウト再現の中心機能とする。

決定順:

1. 印刷範囲
2. 手動改ページ
3. FitToPagesWide / FitToPagesTall
4. 倍率指定
5. 自動改ページ

問題になりうる点:

- 横 1 ページに収める指定と手動改ページが競合する
- 余白と中央寄せを考慮すると見かけの位置が変わる
- 縮小率次第で文字が読めなくなる

他ライブラリの扱い:

- Excel.Report.PDF は余白、中央寄せ、倍率、FitColumn をまとめてスケールへ反映している
- ReoGrid は `PrintableRange`、ユーザー改ページ、自動改ページを分けて管理している
- DioDocs for Excel は印刷設定や図形、チャートを含めた PDF 保存を前提にした製品思想を持つ

OysterReport の方針:

- ReoGrid に近く、印刷範囲、手動改ページ、自動改ページを別概念で持つ
- Excel.Report.PDF に近く、余白、中央寄せ、倍率、横幅合わせを同じスケール決定フェーズで解く
- internal な `PdfRenderPagePlan` にページ確定結果を持つ
- ページ分割理由を保持する
- 既定では可読性を下回る極端な縮小率を警告する

検討が必要な項目:

- Excel の `FitToPagesWide` と手動改ページの競合解決は再現優先順位の定義が必要
- 忠実再現を優先すると、場合によってはページ数増加や可読性低下を招く
- そのため v1 では「Excel の明示設定優先」を基本としつつ、可読性が極端に落ちる場合は診断を出す

### 17.10 今想定している仕様で発生しうる問題

- 列幅換算式の前提が Excel 実機と一致しない
- PDFsharp のフォント計測と Excel 表示結果の差で改行位置がずれる
- 罫線の二重線や斜め罫線で近似表現が必要になる
- 隠し行、隠し列の扱いによって印刷範囲がずれる
- ヘッダ、フッタ、ページ番号と本文領域の競合
- 将来グラフや図形に対応するとレイヤー順とクリップ仕様が不足する

そのため、v1 では以下を仕様に含める。

- 近似を許容する箇所を診断可能にする
- レイアウト確定結果をテストしやすい構造で保持する
- 他ライブラリで忠実度の高い設計は積極的に採用する
- ただし複雑度が高い項目は、段階導入理由を仕様に明記したうえで後続実装へ回す

## 18. 結論

OysterReport の v1 は、Excel を帳票テンプレートとして読み込んで PDF を生成するための「Excel 主体の中間データ構造と PDF 用レンダリング情報を分離したレイアウト変換ライブラリ」として定義する。

この仕様書だけで、Excel をレイアウト編集の正とし、アプリケーションでは限定された帳票操作だけを行い、その結果を `PdfGenerator` の内部 2 フェーズで PDF 化するという全体像を把握できることを目指す。

特に重要なのは以下の 5 点である。

- Excel 依存の読み取り処理と PDF 描画処理を分離すること
- point 基準の中間データ構造を持つこと
- `PdfGenerator` の内部で、中間データ構造とレンダリング情報を 2 フェーズで扱うこと
- フォントと改ページを独立した仕様として明文化すること
- 帳票用途に必要な機能へスコープを絞ること

この方針により、初期実装を過度に複雑化させずに、将来的なテンプレート DSL や高度な帳票機能へ拡張しやすい基盤を作れる。
