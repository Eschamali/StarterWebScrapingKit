---
description: ExecuteCDP / ExecuteBiDi で公式プロトコルを直接呼ぶ低レイヤー操作。同期・非同期、エラー扱い、BiDi+ CDP トンネルまでを解説します。
---

# 低レイヤー BiDi / CDP コマンドについて

高レベル API に無い操作は、公式プロトコルをそのまま呼び出せます。ここが Playwright / Puppeteer における「低レイヤー」と同じ立ち位置です。

## コマンドの単位（ブラウザ / タブ）

コマンドは次の **2 段階**として機能します。

| 単位 | 呼び出し先（例） | 向いている操作 |
| --- | --- | --- |
| **ブラウザ単位** | `CDPBrowser.ExecuteCDP` / `WebDriverBiDiMode.ExecuteBiDi` | 拡張機能の読み込みなど |
| **タブ（コンテキスト）単位** | `CDPContext.ExecuteCDP` / `WebDriverBiDiContext.ExecuteBiDi` | JavaScript 実行・ページ遷移など |

例えば拡張機能は **ブラウザ単位**でしか機能しません。逆に JavaScript 関連は **タブ（コンテキスト）単位**でしか機能しません。

この境界を意識しないと CDP / BiDi からエラーが返ります。ご注意ください。

```vb
' ブラウザ単位（拡張機能など）
t.InheritanceCDPBrowser.ExecuteCDP "Extensions.loadUnpacked", params

' タブ単位（ページ操作など）
t.ExecuteCDP "Page.navigate", params
```

---

## CDP — 同期（`ExecuteCDP`）

コマンドを実行し、結果が返るまで内部で自動待機します。成功時は `result`（Return Object）の中身を [vbacollective-json](https://github.com/vbacollective/json)（`BiDiCDPJson`）として返します。

CDP ドキュメントを見るときの対応関係:

| ドキュメント | 本ツールへの渡し方 |
| --- | --- |
| **Methods** | 第 1 引数にメソッド名文字列をそのまま（例: `"Page.navigate"`） |
| **parameters** | `Dictionary` で Key / Value を組み立てて第 2 引数へ。parameters が無い Methods は省略 |

### 例: `Page.navigate`

[Page.navigate](https://chromedevtools.github.io/devtools-protocol/tot/Page/#method-navigate) の場合:

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext

Dim params As New Dictionary
params.Add "url", "https://example.com"

Dim result As BiDiCDPJson
Set result = t.ExecuteCDP("Page.navigate", params)

' 必要に応じて NodeKey / StringKey などで取り出す
Debug.Print result.StringKey("frameId")
Debug.Print result.StringKey("loaderId")
```

成功時は Return Object を辿って取り出します。詳しい扱いは [vbacollective-json API Reference](https://github.com/vbacollective/json/blob/main/docs/API_REFERENCE.md) にありますが、基本は次を覚えれば十分です。

- `.NodeKey("...")`
- `.StringKey("...")`
- `.NumberKey("...")`
- `.BoolKey("...")`
- `.ExistsKey("...")` / `.Exists("...")`

### エラー扱い（同期）

既定ではエラー時に VBE で停止します。第 3 引数 `StopCDPError:=False` で無視も可能です。

その場合、エラー時は `Nothing` が返るので、これで判定できます（`{ "result": {} }` でもオブジェクト自体は存在するため、`Nothing` と区別できます）。

```vb
Dim result As BiDiCDPJson
Set result = t.InheritanceCDPBrowser.ExecuteCDP("Extensions.loadUnpacked", params, False)

If result Is Nothing Then
    Debug.Print t.InheritanceCDPBrowser.LastCDPJsonError("message")
Else
    Debug.Print result.Stringify
End If
```

参照: [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/)  
デモ: `Demo_CDP.UseExtensions`

---

## CDP — 非同期（`ExecuteCDPAsync`）

コマンド実行後、実行時の **id（`Long`）のみ**を返し、結果は待ちません。

自力で取り出す場合は `TakeEvents` / `TakeResultCDP` を Do ループで呼び出します。同じパターンをまとめた自動待機版 `AutoWaitTakeResultCDP(commandID)` も用意されているので、待つだけでよい場面ではそちらを使うと手早く書けます。

### 向いている使用例

- **クリック後に JavaScript アラートが出て、イベント検知後にアラートを閉じる**  
  → アラート表示中は同期コマンドの結果が返ってこないため
- **複数タブに一斉に `Page.navigate` を投げ、あとでページ読み込みを待つ**  
  → 画面上のタブが一括で変わるレベルで時短できる

```vb
Dim cmdId As Long
cmdId = t.ExecuteCDPAsync("Page.navigate", params)

Do
    t.TakeEvents
    Dim raw As String
    raw = t.TakeResultCDP(cmdId)
    If LenB(raw) Then Exit Do
    DoEvents
Loop
```

::: tip 注意
非同期でも、Pipe / WebSocket 自体にエラーがある場合はエラー停止します。
:::

蓄積件数の上限は `SetLimitCDPResult`（デフォルト 65536）です。上限超過やコマンド ID リセット時は結果履歴がすべて削除されます。

関連: [イベント購読](/guides/events)

---

## BiDi — 同期（`ExecuteBiDi`）

CDP と同様、コマンドを実行して結果が返るまで待機し、成功時は [vbacollective-json](https://github.com/vbacollective/json) として返します。

BiDi ドキュメントを見るときの対応関係:

| ドキュメント | 本ツールへの渡し方 |
| --- | --- |
| **method** | 第 1 引数にメソッド名文字列をそのまま（例: `"browsingContext.navigate"`） |
| **params** | `Dictionary` で組み立てて第 2 引数へ。params が無い場合は省略 |

### 例: `browsingContext.navigate`

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext

Dim params As New Dictionary
params.Add "url", "https://example.com"
params.Add "wait", "complete"
' ※ WebDriverBiDiContext.ExecuteBiDi は内部で context を自動付与します

Dim result As BiDiCDPJson
Set result = t.ExecuteBiDi("browsingContext.navigate", params)
```

結果の取り出しも CDP と同じく `.NodeKey` / `.StringKey` などです。

### エラー扱い（同期）

第 3 引数 `StopBiDiError:=False` で例外にせず `Nothing` 戻りにできます。詳細は `LastBiDiJsonError` を参照。

```vb
Dim params As New Dictionary
Dim extData As New Dictionary
extData.Add "type", "path"
extData.Add "path", "C:\path\to\unpacked-extension"
params.Add "extensionData", extData

Dim result As BiDiCDPJson
Set result = t.ExecuteBiDi("webExtension.install", params, False)

If result Is Nothing Then
    Debug.Print t.InheritanceWebDriverBiDiMode.LastBiDiJsonError("message")
End If
```

参照: [WebDriver BiDi](https://w3c.github.io/webdriver-bidi/)  
デモ: `Demo_WebDriverBiDi.UseExtensions`

::: tip
拡張機能は **ブラウザ（セッション）側**のコマンドです。タブ Context から呼んでも、内部的には Mode 側の権限で動く点に注意してください。
:::

---

## BiDi — 非同期（`ExecuteBiDiAsync`）

CDP の `ExecuteCDPAsync` と同様、実行時 id を返し結果は待ちません。自力で取り出す場合は `TakeEvents` / `TakeResultBiDi` を Do ループで呼び出します。

向いている例も CDP と同じで、ダイアログ待ちや複数コンテキストへの一斉ナビゲートなどです。

Pipe / WebSocket 自体の障害時はエラー停止します。

蓄積件数の上限は `SetLimitBiDi`（デフォルト 65536）です。上限超過やコマンド ID リセット時は結果履歴がすべて削除されます。

---

## BiDi+ CDP トンネル

BiDi に無い細かい CDP を、BiDi 経由で中継します。

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext

Dim params As New Dictionary
Dim result As BiDiCDPJson
Set result = t.ExecuteBiDi("goog:cdp.getSession", params)
Dim sessionId As String
sessionId = result("session")

Set params = New Dictionary
params.Add "method", "Browser.getVersion"
params.Add "params", New Dictionary
params.Add "session", sessionId
Set result = t.ExecuteBiDi("goog:cdp.sendCommand", params)

Debug.Print result.NodeKey("result").StringKey("userAgent")
t.InheritanceWebDriverBiDiMode.quit
```

デモ: `Demo_WebDriverBiDi.TestBiDiPlus_CDPTunnel`

## ConvertToCDPContext

トンネルではなく、同じタブを `CDPContext` として扱う方法です。`CDPElement` 一式が使えます。

```vb
Dim cdp As CDPContext
Set cdp = bidiTab.ConvertToCDPContext
cdp.notify "CDP に変換しました"
```

## 関連

- [`CDPBrowser.ExecuteCDP`](/api/cdp/CDPBrowser#executecdp--executecdpasync)
- [`CDPContext.ExecuteCDP`](/api/cdp/CDPContext#executecdp)
- [`WebDriverBiDiMode.ExecuteBiDi`](/api/bidi/WebDriverBiDiMode)
- [`WebDriverBiDiContext.ExecuteBiDi`](/api/bidi/WebDriverBiDiContext)
- [イベント購読](/guides/events)
