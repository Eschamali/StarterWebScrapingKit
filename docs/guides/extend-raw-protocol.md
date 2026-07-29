# 生プロトコル拡張

高レベル API に無い操作は、公式プロトコルをそのまま呼び出せます。ここが Playwright / Puppeteer における「低レイヤー」と同じ立ち位置です。

## ExecuteCDP

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext

Dim params As New Dictionary
params.Add "path", "C:\path\to\unpacked-extension"

Dim result As BiDiCDPJson
' ブラウザ単位のコマンドは InheritanceCDPBrowser 側
Set result = t.InheritanceCDPBrowser.ExecuteCDP("Extensions.loadUnpacked", params, False)

If result Is Nothing Then
    Debug.Print t.InheritanceCDPBrowser.LastCDPJsonError("message")
Else
    Debug.Print result.Stringify
End If
```

ページ単位なら `t.ExecuteCDP "Network.enable"` のように Context からも呼べます。

参照: [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/)  
デモ: `Demo_CDP.UseExtensions`

`ExecuteCDPAsync` はリクエスト ID（`Long`）を返し、完了はイベント／`TakeEvents` 側で回収する非同期パターンです。

## ExecuteBiDi

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext

Dim params As New Dictionary
Dim extData As New Dictionary
extData.Add "type", "path"
extData.Add "path", "C:\path\to\unpacked-extension"
params.Add "extensionData", extData

Dim result As BiDiCDPJson
Set result = t.ExecuteBiDi("webExtension.install", params, False)
```

参照: [WebDriver BiDi](https://w3c.github.io/webdriver-bidi/)  
デモ: `Demo_WebDriverBiDi.UseExtensions`

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

## エラー扱い

第3引数（CDP）／第3引数（BiDi）の `Stop*Error:=False` で例外にせず `Nothing` 戻りにできます。詳細は `LastCDPJsonError` / `LastBiDiJsonError` を参照。

## 関連

- [`CDPContext.ExecuteCDP`](/api/cdp/CDPContext#executecdp)
- [`WebDriverBiDiContext.ExecuteBiDi`](/api/bidi/WebDriverBiDiContext#executebidi)
- [イベント購読](/guides/events)
