---
description: VBA からは長らく手が届かなかった Excel の一部の機能に、CDP 経由で手が届いた話の入り口です。このセクションの読み方と、なぜ CDP のツールで Excel を扱うのかを紹介します。
---

# ExcelVBAの限界突破

Excel には、年を重ねるごとにいろいろな機能が追加されてきました。

- Power Query
- セルのチェックボックス
- AI (Copilot)
- ...

そして意外にも、それに対応する VBA のメソッドも追加されています。比較的新しいものでいえば、次のようなものです。

| 機能 | 対応する VBA |
| --- | --- |
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

* [Range.CellControl property (Excel)](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cellcontrol)（`Range("A1").CellControl.SetCheckbox` の例）
