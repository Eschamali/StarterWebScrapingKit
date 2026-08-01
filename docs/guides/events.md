# イベント購読

ネットワークやログなどの非同期イベントを VBA 側で受け取ります。デモは `Demo_CDP.ネットワークイベントの確認` / `Demo_WebDriverBiDi.ネットワークイベントの確認` が正本です。

## 2 つの受け取り方

一般的には、[拡張用テンプレート](https://github.com/Eschamali/StarterWebScrapingKit/blob/dev/ForDevelopers/TemplateExtensions/CDP/Normal/exCDP_Template.cls) のように別途 Class オブジェクトを作り、`WithEvents` を定義して対象コンテキスト Class を渡し、非同期イベントごとの処理を書いて…といった儀式が必要です。

一方、`.BrowserEvents`（BiDi なら `.BiDiEvents`）を使う場合は、**標準モジュール（`.bas`）上で直接**イベント処理を行えます。

Demo には `Page.javascriptDialogOpening` を `.bas` 上で処理する例があるので、それを参考に実装してみるとよいでしょう（`Demo_CDP.TestAlert` など）。

::: tip
試作や短いデモでは `BrowserEvents` が手軽です。**長期運用**では、先の Class + `WithEvents` でイベント処理を行うのがベストです。
:::

## CDP（`BrowserEvents`）

1. `BrowserEvents` に `New Dictionary` を渡して記録開始
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

## BiDi（`BiDiEvents`）

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

- [`CDPContext.BrowserEvents` / `SetFilterEvents`](/api/cdp/CDPContext#browserevents--setfilterevents)
- [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [設計思想](/concepts/design-philosophy) — AI トッピング／拡張テンプレート
