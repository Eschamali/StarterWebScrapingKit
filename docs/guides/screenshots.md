# スクリーンショット

CDP の `snapPage` でビューポートまたはフルページを保存します。

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext
t.navigate "https://example.com"

' 第3引数 True でフルページ
t.snapPage Environ("UserProfile") & "\Downloads", "shot.png", False
t.notify "保存しました"

t.InheritanceCDPBrowser.quit
```

デモ: `Demo_CDP.getSnapShot`

BiDi 専用の高レベル API は未掲載です。必要な場合は [`ConvertToCDPContext`](/api/bidi/WebDriverBiDiContext#converttocdpcontext) 後に `snapPage` するか、[低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)で Page.captureScreenshot 相当を呼び出してください。

## 関連

- [`CDPContext.snapPage`](/api/cdp/CDPContext#snappage)
