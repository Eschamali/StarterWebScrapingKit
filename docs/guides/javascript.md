# JavaScript 実行

ページ上で任意の JS を評価します。両プロトコルで `jsEval` があります。

## 基本

::: code-group

```vb [CDP]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")

Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result

' 例外時の戻り値を指定
result = t.jsEval("notDefined.x", , , "fallback")
```

```vb [BiDi]
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext("https://example.com")

Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result
```

:::

## スクリプトの追加（CDP）

```vb
t.jsAddLib "https://cdn.example.com/lib.js"   ' URL
t.jsAddScript "C:\scripts\helper.js"          ' ローカルファイル
```

## 通知バナー（CDP）

```vb
t.notify "処理が完了しました", 5   ' 表示秒数
```

## ダイアログ

アラート処理のデモは CDP / BiDi 双方にあります（`Demo_CDP.TestAlert` / `Demo_WebDriverBiDi.TestAlert`）。BiDi ではセッション capability で自動 dismiss を無効化してからイベント駆動で応答するパターンが典型です。

CDP:

```vb
t.handleDialog True            ' Accept
t.handleDialog False, "入力"   ' prompt 用テキスト
```

## 関連

- [`CDPContext.jsEval`](/api/cdp/CDPContext#jseval)
- [`WebDriverBiDiContext.jsEval`](/api/bidi/WebDriverBiDiContext#jseval)
- [生プロトコル拡張](/guides/extend-raw-protocol)
