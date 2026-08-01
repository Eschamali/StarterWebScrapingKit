# WebDriverBiDiContext

BiDi の browsing context（タブ／ページ）です。Playwright の **Page** に相当します。

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext
t.navigate "https://example.com"
t.InheritanceWebDriverBiDiMode.quit
```

親は `InheritanceWebDriverBiDiMode`（[`WebDriverBiDiMode`](./WebDriverBiDiMode)）。

## ナビ・JS

### `navigate`

```vb
Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
```

### `jsEval`

```vb
Public Function jsEval(JavaScriptStr As String, Optional scriptHandle As String, _
    Optional objectArguments As Collection, Optional IFEXCEPTION As Variant, ...) As Variant
```

[JavaScript 実行](/guides/javascript)

要素のクリック／入力など高レベル API は CDP 側が充実しています。必要なら下記で変換してください。

## CDP への橋渡し

### `ConvertToCDPContext`

```vb
Public Function ConvertToCDPContext() As CDPContext
```

同じタブを [`CDPContext`](/api/cdp/CDPContext) として操作できます。`CDPElement` が使えます。

```vb
Dim cdp As CDPContext
Set cdp = t.ConvertToCDPContext
cdp.getElementByQuery("button").click
```

## プロトコル

### `ExecuteBiDi` / `ExecuteBiDiAsync`

コンテキスト／セッション向け BiDi コマンド。BiDi+（`goog:cdp.*`）もここから。

`ExecuteBiDiAsync` の結果回収・蓄積上限は親の [`TakeResultBiDi`](./WebDriverBiDiMode#takeresultbidi) / [`SetLimitBiDi`](./WebDriverBiDiMode#setlimitbidi) で行います。

```vb
Dim cmdId As Long
cmdId = t.ExecuteBiDiAsync("browsingContext.navigate", params)

Do
    t.InheritanceWebDriverBiDiMode.TakeEvents
    Dim raw As String
    raw = t.InheritanceWebDriverBiDiMode.TakeResultBiDi(cmdId)
    If LenB(raw) Then Exit Do
    DoEvents
Loop
```

[低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)

## 再接続

### `reattach`

```vb
Public Function reattach(userProfile As String, _
    Optional sessionCapabilitiesRequest As Dictionary, _
    Optional WebSocketMode As CDPCoreViaWebSocket) As Boolean
```

最後に操作した BiDi context への再接続。[再接続](/guides/reattach)

### `StartBiDiModeAndConnectTab`

低レベル起動＋タブ接続。

## タイムアウト

`TimeOutSecond` はこのクラス自体にはなく、親の [`WebDriverBiDiMode.TimeOutSecond`](./WebDriverBiDiMode#timeoutsecond) で設定します。

```vb
t.InheritanceWebDriverBiDiMode.TimeOutSecond = 60
```

詳細は [タイムアウト設定方法について](/guides/timeout)。

## 関連

- [`WebDriverBiDiMode`](./WebDriverBiDiMode)
- [CDP と BiDi](/concepts/cdp-vs-bidi)
- [要素の取得](/guides/selectors)（CDP 変換パターン）
- [タイムアウト設定方法について](/guides/timeout)
