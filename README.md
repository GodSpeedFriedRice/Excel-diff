# Excel-diff

2 つの Excel ファイルやフォルダを比較する Windows デスクトップアプリ。WinMerge のように左右でシートを見比べ、値・数式・書式や表示設定の違いを確認する。修正は利用者が Excel で行う。

## サンプルで試す

.NET 10 SDK を入れた Windows で、リポジトリのルートから起動する。

```powershell
dotnet run --project src/ExcelDiff.Desktop
```

1. 「開く…」を押し、「フォルダ比較」を選ぶ。
2. 左に [samples/before](samples/before)、右に [samples/after](samples/after) を指定して「比較」を押す。
3. 結果のファイル行をダブルクリックして詳細を開く。全体プレビューは A1 から始まり、Ctrl+↑ / Ctrl+↓ で差分をたどれる。

サンプルはすべて架空データ。同名の Excel ファイルが左右に7冊ずつあり、見積書の明細追加、工程表の行削除、5,000行目付近の在庫数量変更、納品確認書の印鑑画像などを試せる。

## プレビューの画像サイズ

シート上の PNG・JPEG は、プレビュー用画像の長辺を既定で 512 px に抑え、Excel 上の配置サイズで表示する。必要に応じて、起動前に環境変数 `EXCEL_DIFF_IMAGE_MAX_EDGE` を 1～2048 の整数で設定できる。

```powershell
$env:EXCEL_DIFF_IMAGE_MAX_EDGE = '768'
dotnet run --project src/ExcelDiff.Desktop
```

表示できない貼り付け画像は枠で示す。画像そのものの差分判定は対象外。

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
| `samples/` | フォルダ比較を試すための変更前・変更後の帳票 |

Desktop は入力と結果の表示を担当し、Core は画面に依存せず比較結果を返す。ソリューションとプロジェクトファイルは実装開始時に作成する。
