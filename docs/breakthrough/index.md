---
description: VBA からは長らく手が届かなかった Excel の一部の機能に、CDP 経由で手が届いた話の入り口です。このセクションの読み方と、なぜ CDP のツールで Excel を扱うのかを紹介します。
---

# ExcelVBAの限界突破

Excel には、年を重ねるごとにいろいろな機能が追加されてきました。

- Power Query
- セルのチェックボックス
- AI (Copilot)
- ...

そして意外にも、それに対応する VBA のメソッド/改良も追加されています。比較的新しいものでいえば、次のようなものです。

| 機能 | 対応する VBA |
| --- | --- |
| 正規表現 | `New RegExp`（Microsoft 365 のバージョン 2508 から VBA に組み込み。それまでは VBScript の `VBScript.RegExp` を借りていた） |
| セルのチェックボックス | `Selection.CellControl.SetCheckbox` |
| 計算方法の設定 > 互換性バージョン | `ActiveWorkbook.CompatibilityVersion` |

なんやかんやで、VBA もしっかりとメンテナンスされているんだな～と改めて感じました。

しかし、一部の機能は、いつになっても VBA からはいじれないままでした。このセクションでは、それを扱います。

え、なんで CDP のツールなのに Excel を扱うのかって？

まぁ、読めばわかります。

## このセクションの構成

| ページ | 内容 |
| --- | --- |
| [からくり](/breakthrough/mechanism) | Excel の一部の機能が WebView2 で作られていることと、そこへ CDP で接続できる理由 |
| [Officeスクリプト](/breakthrough/office-scripts) | からくりを使って、VBA から Officeスクリプトを書き込み、実行し、結果を受け取る拡張 |

## 参考資料

* [How to prepare VBA projects for VBScript deprecation](https://devblogs.microsoft.com/microsoft365dev/how-to-prepare-vba-projects-for-vbscript-deprecation/)（Microsoft 365 Developer Blog、2025年9月9日。Windows 版 Office のバージョン 2508（ビルド 19127.20154）から RegExp クラスが VBA に入り、参照設定なしで `New RegExp` と書ける）
* [Range.CellControl property (Excel)](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cellcontrol)（`Range("A1").CellControl.SetCheckbox` の例）
* [Excelの新機能「リストとセル内配列」：1つのセルに複数の値](https://excel-ubara.com/excel1/EXCEL700_04.html)（解説記事内で使ってる例）
