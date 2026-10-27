---
description: ブラウザ起動時のオプションを1つにまとめる設定保持専用クラス BrowserCapabilities。既定値はワークシートから読み込み、Set_○○ で上書きします。
---

# BrowserCapabilities

ブラウザを起動するときの**オプション設定を1つにまとめた、設定保持専用のクラス**です。`OpenBrowser` / `OpenBrowserAndConnectTab` といった「起動系メソッド」の引数に渡します。

起動時のオプションは、通信経路（Pipe / Port / WebView2）、ブラウザの種類（Edge / Chrome × 安定版 / Beta / Dev / Canary）、プロファイル、起動引数、BiDi の設定……と増え続けるため、メソッドの引数で受けるのをやめ、1つのオブジェクトに集約しています。

- 各設定の**既定値は、ワークシート「ブラウザ起動設定」から読み込まれます**（`Class_Initialize`）
- 既定値を変えたい項目だけ、`Set_○○` プロパティで上書きします
- 起動に使ったあとは破棄される **使い切り（ワンショット）** です（後述の「使い切り（ワンショット）の仕様」を参照）

::: tip ヘルパーを使う場合は意識不要です
`ShSetting01_StartBrowser.StartCDPModeContext` などの[起動ヘルパー](/getting-started)は、内部でこのクラスを生成して起動しています。ヘルパーの引数（`StartURL` / `SwitchUser` など）で足りる場合は、このクラスを直接触る必要はありません。
:::

## 使い方

### ワークシートの設定のまま起動する

```vb
Dim t As CDPContext
Set t = New CDPContext
t.OpenBrowserAndConnectTab BrowserCapabilities
```

`BrowserCapabilities` は、そのまま使える**既定インスタンス**です（`New` 不要）。

### 初期 URL を渡して起動する

```vb
BrowserCapabilities.Set_StartURL = "https://example.com"

Dim t As CDPContext
Set t = New CDPContext
t.OpenBrowserAndConnectTab BrowserCapabilities
```

### 一部の設定だけ上書きして起動する

```vb
With BrowserCapabilities
    .Set_UseBrowserPreset = 7                 ' Chrome Canary
    .Set_UserDataDirName = "MyProfile"        ' 別プロファイル
    .Set_AdditionalBrowserArguments = "--disable-popup-blocking"
End With

Dim b As CDPBrowser
Set b = New CDPBrowser
b.OpenBrowser BrowserCapabilities
```

### 自前のインスタンスを使う

既定インスタンスではなく、自分で用意したインスタンスを渡すこともできます。

```vb
Dim caps As BrowserCapabilities
Set caps = New BrowserCapabilities
caps.Set_UseWebSocket = True              ' この起動だけ Port 経由

Dim t As CDPContext
Set t = New CDPContext
t.OpenBrowserAndConnectTab caps
```

### 渡す先のメソッド

| クラス | メソッド | 内容 |
| --- | --- | --- |
| [`CDPContext`](./CDPContext) | `OpenBrowserAndConnectTab(caps)` | ブラウザ起動と、初回タブへの接続まで |
| [`CDPBrowser`](./CDPBrowser) | `OpenBrowser(caps) As String` | ブラウザ起動と接続まで（戻り値は初期 URL） |
| [`WebDriverBiDiContext`](/api/bidi/WebDriverBiDiContext) | `OpenBrowserAndConnectTab(caps)` | BiDi での起動と、初回タブへの接続まで |
| [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode) | `OpenBrowser(caps) As String` | BiDi での起動と接続まで（戻り値は初期 URL） |
| `CDPCoreViaWebSocket` | `AutoConnectDevToolsActivePort(caps) As Boolean` | 起動中のブラウザの `DevToolsActivePort` へ接続（`Set_UserDataDirName` で対象を指定）。[WebSocket モード](/websocket/capabilities) |
| `CDPCoreViaWebView2` | `OpenWebView2AndConnectCDP(caps, Optional AttachHwnd)` | WebView2 の起動と接続。[WebView2 モード](/webview2/capabilities) |

Context 側（`OpenBrowserAndConnectTab`）は、起動直後に確実にスクリプトを実行できるよう、いったん起動スプラッシュ画面を経由してから `Set_StartURL` へ遷移します（[ページ遷移](/guides/navigation)）。この動作は自動で、設定項目はありません。

## 使い切り（ワンショット）の仕様

起動処理の最後に、このオブジェクトは**自動で破棄されます**。渡した変数だけでなく、既定インスタンス（裸の `BrowserCapabilities`）も破棄されるため、次に `BrowserCapabilities` を使うと、**新しいインスタンスとして再生成され、ワークシートの値を読み直します**。つまり、ある起動で上書きした設定が、次の起動へ引き継がれることはありません。

```vb
BrowserCapabilities.Set_StartURL = "https://example.com"
t1.OpenBrowserAndConnectTab BrowserCapabilities    ' ここで破棄される

t2.OpenBrowserAndConnectTab BrowserCapabilities    ' 初期状態（シートの設定）から始まる。StartURL は空
```

## 設定項目（`Set_○○`）

すべて `Property Let`（`Set_BiDiSessionCapabilitiesRequest` のみ `Property Set`）で、同名の `Property Get` で読み出せます。既定値は、ワークシート「ブラウザ起動設定」の対応する項目から読み込まれます。

### ブラウザ起動

| プロパティ | 型 | 既定値の取得元 | 意味 |
| --- | --- | --- | --- |
| `Set_CustomChromiumPath` | `String` | ブラウザexeパス | ブラウザ exe のフルパス（ポータブル版など）。空なら `Set_UseBrowserPreset` に従って自動判定 |
| `Set_UseBrowserPreset` | `Long` | ブラウザタイプ | 使用するブラウザの種類（後述の「`Set_UseBrowserPreset` の値」を参照） |
| `Set_UserDataDirName` | `String` | ユーザーデータフォルダ名 | `--user-data-dir` に付けるフォルダ名。ドライブ付きの絶対パスも指定可能。**接続情報を記録する識別名称**としても使われます |
| `Set_AdditionalBrowserArguments` | `String` | 追加の起動引数 | 追加の起動引数。BiDi で起動するときは、後ろに `--disable-background-timer-throttling` が自動で付きます（後述） |
| `Set_UseWebSocket` | `Boolean` | WebSocketモード | `True` で Port（`--remote-debugging-port`）、`False` で Pipe（`--remote-debugging-pipe`） |
| `Set_StartWindowMode` | `WinState` | 表示ウィンドウモード | 起動時のウィンドウ表示状態（[`WinState`](./CDPContext#winstate)） |
| `Set_StartURL` | `String` | （既定は空） | 起動直後に開く URL |

### WebDriver BiDi

BiDi（[`WebDriverBiDiContext`](/api/bidi/WebDriverBiDiContext) / [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode)）で起動するときだけ使われます。

::: info BiDi で起動するときに自動で付く起動引数
BiDi で起動するときは、`Set_AdditionalBrowserArguments` の後ろに `--disable-background-timer-throttling` を自動で付けます。BiDi-CDP Mapper（`mapperTab.js`）は非表示のタブの中で動くので、裏のタブとしてタイマーを間引かれると、BiDi の処理が遅くなるためです。ChromeDriver も、BiDi のセッションではこの引数を付けてブラウザを起動しています（[Chromium のコミット](https://chromium.googlesource.com/chromium/src/+/f6f5874c8ec33cf232b4fd7aa921fd3ed0e0e7a5)）。
:::

| プロパティ | 型 | 既定値の取得元 | 意味 |
| --- | --- | --- | --- |
| `Set_BiDiViewTab_BiDiCDPMapper` | `Boolean` | シートのデバッグ設定 | BiDi-CDP Mapper のコンソールタブを表示するか |
| `Set_BiDiViewDebugLogForBiDiCDPMapper` | `Boolean` | シートのデバッグ設定 | Mapper 内部のデバッグログを記録するか |
| `Set_BiDiSessionCapabilitiesRequest` | `Dictionary` | （既定は未設定） | `session.new` に渡す `capabilities`。使い方は [はじめに](/getting-started#sessioncapabilitiesrequest-とは) |

```vb
Dim caps As New Dictionary
' ... capabilities を組み立て ...

Set BrowserCapabilities.Set_BiDiSessionCapabilitiesRequest = caps

Dim t As WebDriverBiDiContext
Set t = New WebDriverBiDiContext
t.OpenBrowserAndConnectTab BrowserCapabilities
```

### その他

| プロパティ | 型 | 既定値の取得元 | 意味 |
| --- | --- | --- | --- |
| `Set_EnableUserAgentOverride` | `Boolean` | シートのデバッグ設定 | ブラウザ本来の `User-Agent` の末尾に、このツールのサインを追記するか（[はじめに](/getting-started)） |

### `Set_UseBrowserPreset` の値

| 値 | ブラウザ |
| --- | --- |
| `-1` | プリセット未使用（`Set_CustomChromiumPath` を使う場合） |
| `0` / `1` / `2` / `3` | Edge の安定版 / Beta / Dev / Canary |
| `4` / `5` / `6` / `7` | Chrome の安定版 / Beta / Dev / Canary |

::: tip 内部用のメンバー
`ClsSet_○○` / `ClsGet_○○` で始まるメンバー（`ClsSet_SplashScreenMode`、`ClsGet_CDPCore` など）は、クラス内部の都合による動的な設定用です。ユーザーコードから使う必要はありません。
:::

## 起動ヘルパーとの関係

`ShSetting01_StartBrowser` の4つの起動ヘルパー（`StartCDPModeContext` / `StartCDPMode` / `StartBiDiModeContext` / `StartBiDiMode`）は、内部で次のように `BrowserCapabilities` を使っています。

- `SwitchUser` を指定した場合：新しい `BrowserCapabilities` を生成し、その `Set_UserDataDirName` に `SwitchUser` を設定します
- `SwitchUser` を省略した場合：既定インスタンスを一度破棄して再取得し、ユーザーデータフォルダ名はシートの値を使います
- `StartURL` は、常に `Set_StartURL` へ設定されます（事前に `BrowserCapabilities.Set_StartURL` を設定していても、ヘルパー経由の起動では上書きされます）
- BiDi 版の `sessionCapabilitiesRequest` は、`Set_BiDiSessionCapabilitiesRequest` へ設定されます

## 関連

- [`ReattachOptions`](./ReattachOptions) — 再接続用のオプションオブジェクト
- [`CDPContext`](./CDPContext) / [`CDPBrowser`](./CDPBrowser)
- [`WebDriverBiDiContext`](/api/bidi/WebDriverBiDiContext) / [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode)
- [はじめに](/getting-started)
