# Excel-diff

2 つの Excel ファイルやフォルダを比較する Windows デスクトップアプリ。WinMerge のように左右でシートを見比べ、値・数式・書式や表示設定の違いを確認する。修正は利用者が Excel で行う。

## 設計資料

- [画面設計](docs/ui-design.md)
- [比較仕様](docs/comparison-design.md)

## ディレクトリ構成

| パス | 役割 |
| --- | --- |
| `docs/` | 画面設計と比較仕様 |
| `src/ExcelDiff.Desktop/Views/` | C#／WinForms の比較対象選択、フォルダ比較、Excel 差分詳細 |
| `src/ExcelDiff.Core/Workbooks/` | `.xlsx` の読み込みと比較用データの組み立て |
| `src/ExcelDiff.Core/Comparison/` | ファイルの対応付け、シート・セル・設定の差分判定 |
| `tests/ExcelDiff.Core.Tests/` | 比較処理のテスト |

Desktop は入力と結果の表示を担当し、Core は画面に依存せず比較結果を返す。ソリューションとプロジェクトファイルは実装開始時に作成する。
