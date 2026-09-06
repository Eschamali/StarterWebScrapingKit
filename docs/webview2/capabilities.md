---
description: Excel の UserForm に埋め込んだ WebView2 を、Pipe / WebSocket と同じ CDPContext / CDPElement の感覚で制御する方法を紹介します。
---

# WebView2モードでできること

Pipe・WebSocket が「外部のブラウザプロセス」を相手にするのに対し、WebView2 モードは **Excel自身のUserFormに埋め込んだブラウザ**を制御します。デバッグポートも名前付きパイプも使いません。

- Excel の UserForm に本物の WebView2 を埋め込み、リッチな画面（React / Vue / SPA など）を表示しつつ、同じタブを CDP で操作したい場合
- 社内ツールとして、1枚の xlsm だけで「ブラウザ埋め込みUI」を配布したい場合

詳しい経緯・実装の考え方は [設計思想について](/webview2/design) を参照してください。

## 基本的な接続方法

いちばん簡単なのは、同梱の `WebView2Form` を使う方法です。

```vb
Sub ExcelのユーザーフォームにWebView2を埋め込む()
    With WebView2Form
        If Not .StartCDPModeWebView2 Then Debug.Print "WebView2の初期化に失敗しました。": Exit Sub

        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"
        .show
    End With
End Sub
```

同梱デモ: `Demo_WebView2.ExcelのユーザーフォームにWebView2を埋め込む`

内部では、`WebView2Form.StartCDPModeWebView2` が `CDPCoreViaWebView2.ConnectCDP` を呼んでWebView2の`Environment`/`Controller`/`ICoreWebView2`を生成し、`CDPBrowser.reattachWebView2` / `CDPContext.reattachWebView2` を通じて、Pipe版・WebSocket版と**まったく同じCDPスタック**に接続します。埋め込んでしまえば、`getElementByQuery` や `jsEval` など、これまでのガイドで説明してきた操作がそのまま使えます。

## 自前のUserFormに組み込む場合

`CDPCoreViaWebView2` を直接使えば、自作のUserFormにも組み込めます。

```vb
Public Function ConnectCDP(UserName As String, Optional AttachHwnd As LongPtr) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | WebView2 のユーザーデータフォルダ名 |
| `AttachHwnd` | WebView2 を貼り付けるウィンドウハンドル。省略時は Excel 自身のハンドル（`Application.Hwnd`）を使用 |

```vb
Dim wv2 As New CDPCoreViaWebView2
If Not wv2.ConnectCDP("MyUser", Me.EdgeFrame.hWnd) Then Exit Sub

Dim b As New CDPBrowser
b.reattachWebView2 "MyUser", wv2

Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
```

::: tip 注意
- 既存の接続がある場合は、`ConnectCDP` の再呼び出しで切断・再接続されます
- `CDPBrowser.newTab`（`Target.createTarget`）自体はWebView2モードでも使えますが、WebView2は1インスタンス=1ページのため、新規タブはUserForm内には埋め込まれず**独立した新規ウィンドウ**として開きます。タブ（ウィンドウ）をまたいだCDPコマンドのやり取りは`CallDevToolsProtocolMethodForSession`が担うため、複数の`CDPContext`を並行操作すること自体は可能です（詳細は[設計思想について](/webview2/design)）
:::

## 表示・イベント購読

| メンバー | 役割 |
| --- | --- |
| `Resize(Width, Height, Optional Top, Optional Left)` | WebView2の表示サイズ・位置を変更 |
| `Visible`（Let） | 表示/非表示の切り替え |
| `DevToolsEnabled`（Let） | 右クリックの「検証」等、開発者ツールの有効/無効 |
| `ContextMenuEnabled`（Let） | 右クリックメニューの有効/無効 |
| `SubscribeCdpEvent(EventName) As Boolean` | 指定したCDPイベント名を個別に購読開始 |
| `UnsubscribeCdpEvent(EventName) As Boolean` | 指定したCDPイベント名の購読を解除 |
| `UnsubscribeAllCdpEvents() As Long` | 購読中の全イベントを一括解除（解除件数を返す） |
| `SubscribeCdpEventCount`（Get） | 購読中のイベント数 |
| `isAvailability`（Get） | WebView2（`ICoreWebView2`）が生きているか |

```vb
' Page.loadEventFired を購読してから遷移する例
wv2.SubscribeCdpEvent "Page.loadEventFired"
t.navigate "https://example.com"
' ... TakeEvents ループ等で受信 ...
wv2.UnsubscribeAllCdpEvents
```

::: warning WebSocket/Pipeとの違い
Pipe / WebSocket は「ドメインを`enable`すれば、そのドメインの全イベントが自動で流れてくる」モデルですが、WebView2は`GetDevToolsProtocolEventReceiver`の仕様上、**イベント名ごとの個別購読**が必要です。一括購読の概念はWebView2側に無いため未対応です（一括解除のみ`UnsubscribeAllCdpEvents`として提供）。
:::

## 拡張機能のインストール（v3.1.0〜）

CDPの`Extensions`ドメイン（`Extensions.loadUnpacked`等）は、**WebView2経路だけ`Method not available`で弾かれます**。そのため、拡張機能まわりだけはCDPを介さない専用APIを使います。

![実際に拡張機能をインストールし、UserForm内のWebView2上で動作している様子](/img/拡張機能がWebView2で動作してる様子.png)

*▲ `AddBrowserExtension` でインストールした拡張機能（画像は「Shadowban Scanner」）が、UserForm に埋め込んだ WebView2 上で実際にポップアップを開いて動作している様子*

```vb
Public Function AddBrowserExtension(extensionFolderPath As String) As String   ' 戻り値: インストールした拡張機能のID(失敗時は空文字)
Public Function GetBrowserExtensionIds() As Collection                         ' 各要素はDictionary(キー:"ID"/"Name"/"IsEnabled")
Public Function RemoveBrowserExtension(extensionId As String) As Boolean
```

```vb
With WebView2Form
    '1. 接続前に、拡張機能を有効化しておく(★必須)
    .ThisWebView2.EnvironmentOptions.Set_AreBrowserExtensionsEnabled = True

    '2. WebView2を起動
    If Not .StartCDPModeWebView2 Then Exit Sub
    .show vbModeless

    '3. インストール
    Dim InstallID As String
    InstallID = .ThisWebView2.AddBrowserExtension("C:\path\to\unpacked-extension")
    If LenB(InstallID) = 0 Then MsgBox "拡張機能のインストールに失敗しました": Exit Sub

    '4. 一覧確認
    Dim ext As Variant
    For Each ext In .ThisWebView2.GetBrowserExtensionIds
        Debug.Print ext("ID"), ext("Name"), ext("IsEnabled")
    Next

    '5. アンインストール
    .ThisWebView2.RemoveBrowserExtension InstallID
End With
```

同梱デモ: `Demo_WebView2.拡張機能インストールアンインストール`

::: warning 必ず接続前に有効化する
`AreBrowserExtensionsEnabled` は Environment 生成時にしか読まれない設定です。`ConnectCDP`（`StartCDPModeWebView2`）を呼んだ**あとに** `Set_AreBrowserExtensionsEnabled = True` にしても反映されません。次節の`EnvironmentOptions`と合わせて、**接続前に**設定してください。未設定のままインストールを試みると`ERROR_NOT_SUPPORTED`で失敗します。
:::

## 起動前オプションの設定（`EnvironmentOptions`、v3.1.0〜）

`ICoreWebView2EnvironmentOptions`（`WebView2Loader`がネイティブに提供するはずのオプション）を、VBA側でエミュレーションしたクラスです。`CDPCoreViaWebView2.EnvironmentOptions` から取得し、**`ConnectCDP`を呼ぶ前に**チェーン的に設定します。

```vb
Public Property Get EnvironmentOptions() As WebView2EnvOptions
```

| プロパティ（すべてLet） | 意味 |
| --- | --- |
| `Set_AdditionalBrowserArguments` | ブラウザプロセス起動時の追加コマンドライン引数 |
| `Set_Language` | UI言語 |
| `Set_TargetCompatibleBrowserVersion` | 対象ブラウザバージョン（既定値あり。空文字にすると起動失敗） |
| `Set_AllowSingleSignOnUsingOSPrimaryAccount` | Windowsサインイン中のアカウントでのシングルサインオン可否 |
| `Set_ExclusiveUserDataFolderAccess` | ユーザーデータフォルダの排他アクセス |
| `Set_IsCustomCrashReportingEnabled` | カスタムクラッシュレポートの有効化 |
| `Set_EnableTrackingPrevention` | トラッキング防止の有効化 |
| `Set_AreBrowserExtensionsEnabled` | 拡張機能の使用可否（既定`False`） |
| `Set_ChannelSearchKind` | 探索するEdgeチャンネルの優先順位（`WV2ChannelSearchKind`） |
| `Set_ReleaseChannels` | 探索対象チャンネルのビットマスク（`WV2ReleaseChannels`） |
| `Set_ScrollBarStyle` | スクロールバーの見た目（`WV2ScrollBarStyle`） |

```vb
With WebView2Form.ThisWebView2.EnvironmentOptions
    .Set_AllowSingleSignOnUsingOSPrimaryAccount = False
    .Set_AreBrowserExtensionsEnabled = True
End With
```

同梱デモ: `Demo_WebView2.RunEnvironmentOptionsDemo`

## その他のインターフェースのプロパティ（v3.1.0〜）

拡張機能対応のために`ICoreWebView2` / `ICoreWebView2Controller` / `ICoreWebView2Environment` / `ICoreWebView2Settings` / `ICoreWebView2Profile`のvtableを組み上げたので、ついでにコールバック・イベントを伴わない**スカラー値のプロパティ**は一通り公開しています。用途別に代表例を挙げます（全量はソースコードのコメント、または`Demo_WebView2`内の`Run○○FamilyDemo`各プロシージャを参照）。

| 系統 | 例 | 用途 |
| --- | --- | --- |
| `ICoreWebView2Controller` | `ZoomFactor`（Let）/ `RasterizationScale`（Let）/ `SetBoundsAndZoomFactor` / `MoveFocus` / `SetDefaultBackgroundColor` | 表示倍率・DPI・フォーカス・背景色 |
| `ICoreWebView2` | `IsMuted`（Let）/ `IsDocumentPlayingAudio`（Get）/ `StatusBarText`（Get）/ `FaviconUri`（Get） | 音声ミュート、再生中判定、ステータスバー、favicon |
| `ICoreWebView2Environment` | `userDataFolder`（Get）/ `FailureReportFolderPath`（Get） | 実際に使われているフォルダパスの確認 |
| `ICoreWebView2Settings` | `ScriptEnabled` / `WebMessageEnabled` / `UserAgentOverride` / `GeneralAutofillEnabled`（いずれもLet） | JS実行・WebMessage・UA偽装・オートフィルの可否 |
| `ICoreWebView2Profile` | `ProfileName`（Get）/ `IsInPrivateModeEnabled`（Get）/ `DefaultDownloadFolderPath`（Let）/ `PreferredTrackingPreventionLevel`（Let） | プロファイル情報、ダウンロード先、トラッキング防止レベル |

::: info あえて実装していないもの
「あくまでもCDP制御主体のツール」という線引きのため、次の3種類は対象外です。
- 完了コールバックが必要なメソッド（`ExecuteScriptAsync`等）
- WebView2ネイティブのイベント（`NavigationCompleted`等）— CDPの`Page.*`イベントで代替してください
- 他のインターフェースへネストして依存するもの
:::

## 再接続 (reattach)

Pipe / WebSocket と同じく、`reattachWebView2` で既存のWebView2接続情報に再接続できます。

```vb
Public Function reattachWebView2(userProfile As String, WebView2Mode As CDPCoreViaWebView2, Optional reuseSession As Boolean) As Boolean
```

詳細は [再接続 (reattach)](/guides/reattach) を参照してください。

## 関連

- [設計思想について](/webview2/design) — 機械語サンク・vtable、移植元へのクレジット
- [Excel単独で「真のWebView2」を完全制御する](/userform/vba-only) — UserForm埋め込みの詳しい解説
- [再接続 (reattach)](/guides/reattach)
- デモ: `Demo_WebView2.ExcelのユーザーフォームにWebView2を埋め込む`
