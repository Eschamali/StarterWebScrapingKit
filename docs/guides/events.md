# イベント購読

ネットワークやログなどの非同期イベントを VBA 側で受け取ります。デモは `Demo_CDP.ネットワークイベントの確認` / `Demo_WebDriverBiDi.ネットワークイベントの確認` が正本です。

## CDP

1. `BrowserEvents` に `New Dictionary` を渡して記録開始
2. 必要に応じて、`○○.enable` でイベントを有効化
2. 必要なら `SetFilterEvents` でイベント名を絞る
3. ドメインを `ExecuteCDP` で enable
4. 操作後、Dictionary を読む（必要ならファイルへ保存）

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext

t.SetFilterEvents = "Network.requestWillBeSent"
t.SetFilterEvents = "Network.loadingFinished"
Set t.BrowserEvents = New Dictionary

t.ExecuteCDP "Network.enable"
t.navigate "https://example.com"

' t.BrowserEvents に蓄積
' 記録を止める: Set t.BrowserEvents = Nothing
' セーブデータを戻して再開も可能

t.InheritanceCDPBrowser.quit
```

フィルタ未設定時はキャプチャ対象が広くなります。本番では必要なイベントだけに絞ってください。

## BiDi

1. `BiDiEvents` に `New Dictionary`
2. `sessionSubscribe` にイベント名の `Collection` を渡す
3. 操作後 `TakeEvents` で受信キューを吸い上げる

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext

Set t.InheritanceWebDriverBiDiMode.BiDiEvents = New Dictionary

Dim events As New Collection
events.Add "network.beforeRequestSent"
events.Add "network.responseCompleted"
events.Add "log.entryAdded"
Set t.InheritanceWebDriverBiDiMode.sessionSubscribe = events

t.navigate "https://example.com"
t.InheritanceWebDriverBiDiMode.TakeEvents

' t.InheritanceWebDriverBiDiMode.BiDiEvents を参照

t.InheritanceWebDriverBiDiMode.quit
```

## セーブ／再開

両デモとも、蓄積 Dictionary を別変数に退避 → `Nothing` で停止 → 再代入で再開、というパターンを示しています。長時間ジョブで「区間だけ聞きたい」ときに便利です。

## 関連

- [`CDPContext`](/api/cdp/CDPContext)
- [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode)
- [生プロトコル拡張](/guides/extend-raw-protocol)
