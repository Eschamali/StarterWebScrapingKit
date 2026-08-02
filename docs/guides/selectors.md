---
description: CDPElement による要素取得。ID / CSS / XPath、存在待ち、Shadow DOM・iframe、BiDi から CDP へ変換して操作する方法を解説します。
---

# 要素の取得

CDP では `CDPElement` が中心です。BiDi 側に同等の高レベル要素 API はまだ薄いので、**要素操作が必要なら CDP**、または `ConvertToCDPContext` を使います。

## セレクタの種類（CDP）

| メソッド | 説明 |
| --- | --- |
| `getElementByID` | `id` 属性 |
| `getElementByQuery` / `getElementsByQuery` | CSS セレクタ |
| `getElementByXPath` / `getElementsByXPath` | XPath |

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")

Dim el As CDPElement
Set el = t.getElementByID("submit")
Set el = t.getElementByQuery("form input[name='q']")
Set el = t.getElementByXPath("//button[contains(.,'送信')]")

Dim list As Collection
Set list = t.getElementsByQuery("a.item")
```

## 存在待ち

```vb
' 最大 30 秒待ってから操作
t.getElementByID("ready").onExist.click

' 存在確認のみ
If t.getElementByQuery(".toast").isExist Then
    Debug.Print "shown"
End If
```

## Shadow DOM / iframe（CDP）

```vb
Dim host As CDPElement
Set host = t.getElementByQuery("my-widget")
Dim root As CDPElement
Set root = host.GetShadowRoot
root.getElementByQuery("button").click

Dim frame As CDPElement
Set frame = t.getElementByQuery("iframe#app").getIFrame
frame.getElementByID("inner").click
```

::: tip 💡TIP
Shadow DOM内の操作イメージは、Shadow DOM手前まで要素を取得し、`GetShadowRoot`で向こうに渡るイメージとなります。  
所謂、通り抜けフープ的な概念です。  
状況に応じて「何番目のShadowRoot」ということも可能です。
:::

デモ: `Demo_CDP.SimpleShadowRootTest` / `iframeShadowRootTest` / `runIFrame`

## 日本語 id（CDP）

`id` に日本語が含まれるページでは、設定シートの **常に UTF-8 で CDP-Json 送信** を ON にしてください。  
デモ: `Demo_CDP.JapaneseElementTest`
![UTF-8のスイッチング](/img/JPSend.png)

## BiDi からの要素操作

```vb
Dim bidi As WebDriverBiDiContext
Set bidi = ShSetting01_StartBrowser.StartBiDiModeContext("https://example.com")

Dim cdp As CDPContext
Set cdp = bidi.ConvertToCDPContext
cdp.getElementByQuery("button").click
```

::: tip 💡TIP
または [JavaScript 実行](/guides/javascript) の `jsEval` で DOM を直接操作します。
:::

## 次へ

- [入力とクリック](/guides/input)
- [`CDPElement` API](/api/cdp/CDPElement)
