---
description: CDP の非同期イベント（{"method":...}）を VBA でどう受け取るか。登録制コールバック・Select Case 固定・RaiseEvent という3つの方式を、拡張性の観点から比較します。
---

# イベント処理と拡張性

CDP から流れてくるメッセージには2種類あります。コマンドの応答（`{"id":..., "result":...}`）と、ブラウザが勝手に送ってくる非同期イベント（`{"method":..., "params":...}`）です。後者をどう利用者に届けるかは、ライブラリの性格が最も出るところです。

3プロジェクトは、ここで**まったく違う3つの答え**を選んでいました。

## 方式の比較

| | 配送方式 | モデル |
| --- | --- | --- |
| **StarterWebScrapingKit** | VBA ネイティブの `RaiseEvent` × 4種 | イベント駆動（pub/sub） |
| **VBAChromeDevProtocol** | イベント名をキーにした登録制コールバック | 中央集権コールバック |
| **vba-cdp-webdriver** | `Select Case` で自身の状態を直接更新 | ポーリング型の状態キャッシュ |

## StarterWebScrapingKit ―― `RaiseEvent` 4種

`CDPCore.cls` は4つの正式な VBA イベントを宣言しています。

```vb
Public Event CDPBrowserEvent(methodName As String, RawJson As String)                     'ブラウザに対する非同期イベント
Public Event CDPBrowserID(id As Long, RawJson As String)                                  'ブラウザに対するCDPコマンド結果
Public Event CDPContextEvent(methodName As String, RawJson As String, sessionID As String) 'タブに対する非同期イベント
Public Event CDPContextID(id As Long, RawJson As String, sessionID As String)              'タブに対するCDPコマンド結果
```

`method` があるか / `id` があるか、そして `sessionId` があるかないか。CDP のプロトコル構造そのものが、そのままイベントのシグネチャの違いとして表現されています。受け取る側は `WithEvents` で購読するだけです。

```vb
Private WithEvents ex_CDPCore As CDPCore

Private Sub ex_CDPCore_CDPContextEvent(methodName As String, RawJson As String, sessionID As String)
    If CurrentSessionId <> sessionID Then Exit Sub   ' このタブ以外は捨てる

    Select Case methodName
        Case "お好きなmethod名"
            ' ここに自分の処理を書くだけ
    End Select
End Sub
```

これは `ForDevelopers/TemplateExtensions/CDP/Normal/exCDP_Template.cls` の雛形そのままです。**コア本体（`CDPCore.cls`）を一切編集せず**、新しいクラスファイルを1つインポートするだけで拡張が成立します。

## VBAChromeDevProtocol ―― 登録制コールバック

`clsCDP.cls` は、興味のあるイベント名だけを辞書に登録させる方式です。

```vb
' registers an object to handle events as received as opposed to processing message queue after response
' returns current handler if one exists, set to Nothing to remove handler
Public Function registerEventHandler(ByVal eventName As String, eventHandler As Object) As Object
    If eventHandlers.Exists(eventName) Then Set registerEventHandler = eventHandlers(eventName)
    If eventHandler Is Nothing Then
        If eventHandlers.Exists(eventName) Then eventHandlers.Remove eventName
    Else
        Set eventHandlers(eventName) = eventHandler
    End If
End Function
```

使う側はこうなります。

```vb
Dim dlProgressBar As New ehDownloadProgress
browser.cdp.registerEventHandler "Page.downloadWillBegin", dlProgressBar
```

```vb
' ehDownloadProgress.cls （新規クラス、コア無改造）
Public Function processEvent(ByVal eventName As String, ByVal eventData As Dictionary) As Boolean
    If eventName = "Page.downloadWillBegin" Then
        ' 自分の処理
        processEvent = True   ' True を返せばキューに積まれない
    End If
End Function
```

**こちらもコア無改造で拡張できます。** 設計の地力としては互角で、しかも実行時に動的に登録・解除できる（`Nothing` を渡せば外れる、戻り値で旧ハンドラを受け取れる）という、`WithEvents` にはない柔軟さもあります。

差が出るのは細部です。

### 1イベント名につき1ハンドラ

宣言部にはっきり書かれています。

```vb
' registered event handlers
' key is event name, only 1 event handler allowed per event name
Private eventHandlers As Dictionary
```

`Dictionary` なので、同じイベント名に2つ目を登録すると1つ目が黙って上書きされます。戻り値で旧ハンドラを受け取れるため手動でチェーンは組めますが、自動ではありません。**複数の拡張機能が同時に `Page.frameNavigated` を見たい**というケースでは、自分で連鎖処理を書く必要があります。

`WithEvents` の側はここが逆で、同じ `CDPCore` インスタンスに何個でも独立した購読者を貼れます。拡張クラスAとBが両方同じイベントを見ても、互いに干渉しません。

### バインディングが命名規約ベース

`eventHandler As Object` に対する `eventHandler.processEvent(...)` は遅延バインディングです。VBA の `Implements` を使っていないので、**メソッド名や引数を間違えてもコンパイル時には検知されず、そのイベントが実際に発火した瞬間に初めて実行時エラー**になります。`WithEvents` の側はシグネチャが合わなければコンパイルが通りません。

### 拾わなかったイベントは、次の送信で消える

登録されていない（あるいは `processEvent` が `False` を返した）イベントは内部キューに積まれますが、次にコマンドを1つ送るとキューごと破棄されます。

```vb
' Before sending a message the messagebuffer is emptied
' All messages that we have received sofar cannot be an answer
' to the message that we will send
' So they can be safely discarded
clearMessageQueue ' discard any messages not already processed
```

`registerEventHandler` を使っている限り受信の瞬間にディスパッチされるので実害はありませんが、`getMessageQueue` を直接読む裏ルートに頼るとコマンド送信のタイミング次第で消えます。

### 常駐監視の既製品がない

`peakMessage` は `Public Function` なので、自分でループを回してイベントだけ監視し続けることは技術的に可能です。ただしそれを手軽にやるためのタイマー統合機構は同梱されていません。StarterWebScrapingKit 側は `Advanced/exCDP_TemplateWithSafeTimer.cls` として VBA-SafeTimer 統合済みの雛形があり、`StartCheckAsyncEvents 50` の一行で 50ms 間隔の常駐監視が立ち上がります。ここは「できるか / できないか」ではなく**車輪を自分で作るか、完成品があるか**の差です。

## vba-cdp-webdriver ―― `Select Case` に決め打ち

3つの中で最もシンプルで、最も閉じています。`a5_CDPEventHandler.cls` の `GetInfo` という1本の Sub がメッセージポンプから直接呼ばれ、既知のイベントだけを `Select Case` で解釈して、自身の Public なフラグやコレクションを書き換えます。

```vb
Public Sub GetInfo(EventInfo As String)
    ' ...
    Select Case eventName
        Case "Page.javascriptDialogOpening"
        Case "Page.fileChooserOpened"
        Case "Target.targetCreated"
        Case "Network.requestWillBeSent"
        Case "Page.downloadProgress"
        ' ... 全22ケース
    End Select
End Sub
```

呼び出し側は `events.NetworkInFlight` や `events.DialogInfoDic("IsExistDialog")` のように、**後からポーリングして状態を読む**だけです。外部への通知・イベント発行は一切ありません。

新しい `method` を拾いたければ、この `Select Case` 自体にコードを追記する ―― つまり**ライブラリのコア本体を改造する**しかありません。しかも上位の `IWebDriver` / `ChromeDriver` が生の CDP メッセージを隠蔽する設計なので、自分のコードから `{"method":...}` の生データを覗く手段もありません。

::: info 「閉じている」ことが常に悪いわけではない
この設計は拡張性を犠牲にする代わりに、**よくある非同期イベント処理が最初から完成品として載っている**という利点を得ています。`ClickAndThenAlertDialogErase`（クリックしたら出るアラートを自動で消す）、`DownloadWatchStart`、`SetInterceptFileChooserDialog`、`EnableNetworkInterception` / `AddBlockedURLPattern` ―― 22ケースぶんのイベント処理が、すでに使える形で API 化されています。自分でイベントハンドラを書く場面自体が最初から少ないなら、この割り切りは合理的です。
:::

## まとめ

| 観点 | StarterWebScrapingKit | VBAChromeDevProtocol | vba-cdp-webdriver |
| --- | --- | --- | --- |
| コア無改造での拡張 | できる | できる | できない |
| 複数ハンドラの共存 | 何個でも独立購読 | 同名イベントは上書き | — |
| バインディング | 型付きイベント（コンパイル時検知） | 命名規約（実行時エラー） | — |
| 未処理イベント | 全て発火、破棄なし | 次のコマンド送信で消える | 無視される |
| セッション情報 | 引数として渡される | 生 JSON から自分で抽出 | 扱わない |
| 常駐監視 | SafeTimer 統合テンプレ同梱 | 自作が必要 | — |
| ドキュメント | 公式テンプレ + README | サンプル1個 | なし |

拡張性という一点では **StarterWebScrapingKit と VBAChromeDevProtocol はほぼ互角**で、差は「テンプレート化・文書化されているか」と「型で守られているか」に集約されます。1つの拡張を自分だけで使うなら VBAChromeDevProtocol でも十分ですが、複数の拡張を同時に動かす・長期的にメンテするという条件が付くと、`WithEvents` の側に分があります。

## 関連

- [ディスパッチと購読モデル](/core-comparison/dispatch) — Puppeteer / Playwright の `EventEmitter` との比較
- [マルチタブとセッション管理](/vba-comparison/multi-tab) — `sessionId` の扱いがイベント配送にどう効くか
