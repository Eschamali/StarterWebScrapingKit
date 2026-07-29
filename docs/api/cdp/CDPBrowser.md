# CDPBrowser

ブラウザプロセス単位のエントリです。Playwright の **Browser** に相当します。

通常は `ShSetting01_StartBrowser.StartCDPMode` で取得します。タブ操作だけなら [`CDPContext`](./CDPContext) を返す `StartCDPModeContext` の方が簡単です。

```vb
Dim b As CDPBrowser
Set b = ShSetting01_StartBrowser.StartCDPMode
Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
b.quit
```

## 起動・再接続・終了

### `start`

```vb
Public Sub start(Optional Name As String = "chrome", ...)
```

ブラウザを起動しパイプ接続します。日常利用では設定シート経由を推奨。

### `reattach`

```vb
Public Function reattach(userProfile As String, Optional WebSocketMode As CDPCoreViaWebSocket) As Boolean
```

既存セッションへ再接続。[再接続ガイド](/guides/reattach)

### `quit`

```vb
Public Sub quit()
```

ブラウザを終了しリソースを解放します。

### `isLiveBrowser`

```vb
Public Function isLiveBrowser() As Boolean
```

プロセス／接続が生きているか。

## タブ

### `newTab`

```vb
Public Function newTab(Optional Url As String, Optional newWindow As Boolean, _
    Optional setMain As Boolean, Optional isHidden As Boolean, _
    Optional browserContextId As String, Optional isBackground As Boolean) As CDPContext
```

新規タブ（またはウィンドウ）を開き [`CDPContext`](./CDPContext) を返します。

### `getTab`

```vb
Public Function getTab(Optional tabName As String, Optional Url As String, _
    Optional setMain As Boolean, Optional SearchTypeID As TargetgetTargetsType, _
    Optional doRetrySecond As Double) As CDPContext
```

既存タブを検索して接続。reattach 後は `setMain:=True` を推奨。

### `TabCount`

```vb
Public Function TabCount() As Long
```

### `attachToTab` / `DiscardSessionID`

低レベルなセッション管理用。

## プロトコル

### `ExecuteCDP` / `ExecuteCDPAsync`

```vb
Public Function ExecuteCDP(methodName As String, _
    Optional params As Scripting.Dictionary, _
    Optional StopCDPError As Boolean = True) As BiDiCDPJson

Public Function ExecuteCDPAsync(...) As Long
```

ブラウザターゲット向け CDP コマンド。[生プロトコル拡張](/guides/extend-raw-protocol)

### `TakeEvents`

非同期応答／イベントの吸い上げ。

### `LastCDPJsonError`

直前エラー（Dictionary 風アクセス）。`StopCDPError:=False` 時に参照。

## その他

| メンバー | 説明 |
| --- | --- |
| `openDevTools` | 指定ターゲットで DevTools を開く |
| `printTargetInfos` / `printParams` | デバッグ出力 |
| `sleep` | 秒待ち |
| `TimerCounter` | 経過時間 |
| `serializeForMainTab` | メインタブの session/target を記録 |

## 関連

- [マルチタブ](/guides/multi-tab)
- [`CDPContext`](./CDPContext)
