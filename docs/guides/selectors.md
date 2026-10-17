---
description: CDPElement / WebDriverBiDiElement による要素取得。ID / CSS / XPath、存在待ち、Shadow DOM・iframe、BiDi での要素操作を解説します。
---

# 要素の取得

CDP では `CDPElement`、WebDriver BiDi では `WebDriverBiDiElement` が中心です。どちらも、**呼び出すメソッドはほぼ同じ**です。以下は CDP の例で説明し、BiDi の違いは末尾の「BiDi からの要素操作」の節でまとめます。

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

' Shadow Root の中は XPath でも探せます（先頭に `.` を付けた相対指定）
root.getElementByXPath(".//button[@id='ok']").click

' 同一オリジンの iframe 要素から入る
Dim frame As CDPElement
Set frame = t.getElementByQuery("iframe#app").getIFrame
frame.getElementByID("inner").click
```

Shadow Root 内で XPath を使うときは、`./○○`（直下）や `.//○○`（配下の子孫）のように先頭に `.` を付けてください。子要素が1つもない Shadow DOM では XPath が使えないため、CSS セレクタで探します（詳細は [`CDPElement.getElementByXPath`](/api/cdp/CDPElement#getelementbyxpath--getelementsbyxpath)）。

名前 / URL で iframe を探して `executionContextId` 経由で入る場合は、[`CDPContext.printChildFrames`](/api/cdp/CDPContext#printchildframes) → [`getIFrameContextID`](/api/cdp/CDPContext#getiframecontextid) → [`getIFrame`](/api/cdp/CDPContext#getiframe) です。

```vb
t.printChildFrames   ' 直下の子 iframe をログ確認

Dim ctxId As Long
ctxId = t.getIFrameContextID(iframeName:="app-frame")

Dim frame As CDPElement
Set frame = t.getIFrame(ctxId)
frame.getElementByQuery("button").click
```

::: tip 💡TIP
Shadow DOM内の操作イメージは、Shadow DOM手前まで要素を取得し、`GetShadowRoot`で向こうに渡るイメージとなります。  
所謂、通り抜けフープ的な概念です。  
状況に応じて「何番目のShadowRoot」ということも可能です。
:::

デモ: `Demo_CDP.SimpleShadowRootTest` / `iframeShadowRootTest` / `runIFrame`

## 日本語 id（CDP）

`id` に日本語が含まれるページでも、そのまま指定できます（CDP-Json は常に UTF-8 で送受信します）。  
デモ: `Demo_CDP.JapaneseElementTest`

## BiDi からの要素操作

`WebDriverBiDiContext` の `getElementBy*` で、[`WebDriverBiDiElement`](/api/bidi/WebDriverBiDiElement) を取得します。使い方は `CDPElement` と同じです。

```vb
Dim bidi As WebDriverBiDiContext
Set bidi = ShSetting01_StartBrowser.StartBiDiModeContext("https://example.com")

bidi.getElementByQuery("input[name='q']").value = "検索語"
bidi.getElementByXPath("//button[text()='検索']").onExist.click
```

CDP との違いは次の2点です。

| 操作 | CDP | BiDi |
| --- | --- | --- |
| iframe | `CDPElement.getIFrame` | iframe を [`WebDriverBiDiMode.getTab`](/api/bidi/WebDriverBiDiMode) の `maxDepth:=1` で取得してから検索 |
| Shadow DOM | `GetShadowRoot(Index)` / `GetShadowRoots` | `GetShadowRoot()` で1段ずつ辿る（複数取得はなし） |

```vb
' BiDi: iframe 内の要素
Dim frame As WebDriverBiDiContext
Set frame = bidi.ThisWebDriverBiDiMode.getTab("https://example.com/frame.html", maxDepth:=1)
frame.getElementByQuery("#inner").click

' BiDi: Shadow DOM を1段ずつ
bidi.getElementByQuery("my-widget").GetShadowRoot.getElementByQuery("button").click
```

CDP 独自の機能（`getIFrame` など）を使いたい場合は、[`UpgradeBiDiPlus`](/api/bidi/WebDriverBiDiContext#upgradebidiplus) で同じタブを `CDPContext` に変換して、`CDPElement` で操作することもできます。

```vb
Dim cdp As CDPContext
Set cdp = bidi.UpgradeBiDiPlus
cdp.getElementByQuery("button").click
```

## 次へ

- [入力とクリック](/guides/input)
- [`CDPElement` API](/api/cdp/CDPElement) / [`WebDriverBiDiElement` API](/api/bidi/WebDriverBiDiElement)
