---
description: ネットワークやログなどの非同期イベントを VBA で受け取る方法。BrowserEvents / BiDiEvents と WithEvents 拡張の使い分けを解説します。
---

# イベント購読

ネットワークやログなどの非同期イベントを VBA 側で受け取ります。デモは `Demo_CDP.checkNetworkEvents` / `Demo_WebDriverBiDi.checkNetworkEvents` が正本です。

## 2 つの受け取り方

一般的には、[拡張用テンプレート](https://github.com/Eschamali/StarterWebScrapingKit/blob/dev-xlflow/ForDevelopers/TemplateExtensions/CDP/Normal/exCDP_Template.cls) のように別途 Class オブジェクトを作り、`WithEvents` を定義して対象コンテキスト Class を渡し、非同期イベントごとの処理を書いて…といった儀式が必要です。

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

t.ThisCDPBrowser.CloseBrowser
```

フィルタ未設定時はキャプチャ対象が広くなります。本番では必要なイベントだけに絞ってください。

## BiDi（`BiDiEvents`）

1. `BiDiEvents` に `New Dictionary`
2. `SubscribeBiDiEvent` にイベント名を渡す（`Collection` または 1次元配列）
3. 操作後 `TakeEvents` で受信キューを吸い上げる

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext

Set t.ThisWebDriverBiDiMode.BiDiEvents = New Dictionary
t.ThisWebDriverBiDiMode.SubscribeBiDiEvent = Array("network.beforeRequestSent", "network.responseCompleted", "log.entryAdded")

t.navigate "https://example.com"
t.ThisWebDriverBiDiMode.TakeEvents

' t.ThisWebDriverBiDiMode.BiDiEvents を参照

t.ThisWebDriverBiDiMode.CloseBrowser
```

::: tip タブ（コンテキスト）単位での購読（v3.1.1.1〜）
上記はブラウザ全体（全タブ共通）での購読です。タブAはネットワーク系イベントを購読するが、タブBは何も購読しない、といった使い分けをしたい場合は、`ThisWebDriverBiDiMode.SubscribeBiDiEvent` の代わりに `t.SubscribeBiDiEvent`（[`WebDriverBiDiContext.SubscribeBiDiEvent`](/api/bidi/WebDriverBiDiContext#subscribebidievent)、v3.2.0で`sessionSubscribe`から改称）を使ってください。不要なイベントを流さない分、パフォーマンス低下を防げます。
:::

## 拡張クラス（`WithEvents`）で受け取る

長期運用向けの書き方です。拡張用のクラスを作り、`CDPContext`（タブ）や `CDPBrowser`（ブラウザ）を `WithEvents` で保持して、`ExEvent` イベントで受け取ります。

```vb
' 拡張用クラス（例: exCDP_MyWatcher.cls）
Private WithEvents ex_CDPContext As CDPContext
Private WithEvents ex_CDPBrowser As CDPBrowser    ' ブラウザ単位のイベントが必要な場合のみ

Friend Sub Init(Inheritance As CDPContext)
    Set ex_CDPContext = Inheritance
    Set ex_CDPBrowser = Inheritance.ThisCDPBrowser
End Sub

' このタブ宛ての非同期イベント
Private Sub ex_CDPContext_ExEvent(methodName As String, RawJson As String)
    Select Case methodName
        Case "Page.javascriptDialogOpening"
            ' ここに処理を書く
    End Select
End Sub

' ブラウザ単位の非同期イベント（Target.targetCreated など）
Private Sub ex_CDPBrowser_ExEvent(methodName As String, RawJson As String)
End Sub
```

- `CDPContext` の `ExEvent` には、そのタブ宛てのイベントだけが届きます
- BiDi では、[`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode) の `ExEvent` で受け取ります（`WebDriverBiDiContext` にはイベントがありません）
- イベントが届くのは、`TakeEvents` などで受信処理が回ったときです

## セーブ／再開

両デモとも、蓄積 Dictionary を別変数に退避 → `Nothing` で停止 → 再代入で再開、というパターンを示しています。長時間ジョブで「区間だけ聞きたい」ときに便利です。

## 関連

- [`CDPContext.BrowserEvents` / `SetFilterEvents`](/api/cdp/CDPContext#browserevents--setfilterevents)
- [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [設計思想](/concepts/design-philosophy) — AI トッピング／拡張テンプレート
