---
description: ブラウザ（ページの JavaScript）と VBA のあいだで、データや合図を両方向にやり取りする方法。IE 時代の方法と比べながら、このツールで使える 5 つの選択肢と、それぞれの向き・不向きを解説します。
---

# ブラウザと VBA の双方向やり取り

## はじめに

IE（`InternetExplorer.Application` や、UserForm に置く WebBrowser コントロール）を VBA で操作していたころは、VBA とWebページのあいだで、両方向にやり取りできました。参照設定に「Microsoft Internet Controls」と「Microsoft HTML Object Library」を足すだけで、次のようなことができました。

### VBA からページへ

| 手段 | 内容 |
| --- | --- |
| `Document` の DOM | `ie.Document.getElementById("q").Value = "..."` のように、ページの DOM を VBA から直接読み書きする |
| `Document.parentWindow.execScript` | JavaScript の文字列を、ページの中で実行する。IE11 で廃止され、`eval` に置き換えられた |
| `Document.Script` | ページのスクリプトエンジンの `IDispatch` を受け取る。ページの JavaScript の関数を、VBA から呼ぶのに使われた |

### ページから VBA へ

| 手段 | 内容 |
| --- | --- |
| `WithEvents` で `InternetExplorer` / WebBrowser を持つ | `DocumentComplete`、`BeforeNavigate2`、`NewWindow2`、`OnQuit` などのイベント（`DWebBrowserEvents2`）を、VBA のイベントプロシージャで受け取る |
| `WithEvents` で `HTMLDocument` を持つ | 読み込み完了（`DocumentComplete`）のあとに `Document` を受け取り、`doc_onclick` などで、ページ上のクリックなどを受け取る |
| 要素の `onclick` に VBA のオブジェクトを代入する | 既定のメソッドを持つクラスのオブジェクトを `onclick` に代入すると、クリックされたときに、そのメソッドが呼ばれる。ボタンのイベントは WebBrowser コントロールまで伝わらない（バブリングしない）ため、要素ごとにこうして受け取った |
| `window.external` | ホストが自分の `IDispatch` を渡すと、ページの JavaScript から呼べる（.NET の WebBrowser では `ObjectForScripting`）。VBA から用意するのは難しい |

たとえば、ページ上のクリックは次のように受け取れました（クラスモジュール内）。

```vb
Private WithEvents ie As SHDocVw.InternetExplorer
Private WithEvents doc As MSHTML.HTMLDocument

Private Sub ie_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Set doc = ie.Document          ' 読み込みが終わってから受け取る
End Sub

Private Function doc_onclick() As Boolean
    Debug.Print "クリックされました"
    doc_onclick = True             ' IE 側の既定の処理を続ける
End Function
```

IE は、VBA と同じ COM の世界にいました。そのため、ページの `Document` や `window` を VBA のオブジェクトとして受け取り、イベントも `WithEvents` でそのまま受け取れたわけです。

一方、このツールが扱う Edge / Chrome は、VBA とは別のプロセスで動いています。VBA の手元にあるのは、CDP（または WebDriver BiDi）でやり取りする JSON のメッセージだけです。そのため、IE のようにページのオブジェクトを直接つかむことはできません。そのかわり、次の選択肢があります。

## このツールでの選択肢

| 方式 | 向き | 戻り値 | 大容量 | 追加のサンク | 通常の Edge | 評価 |
| --- | --- | --- | --- | --- | --- | --- |
| `Runtime.evaluate` / `Runtime.callFunctionOn` | VBA→JS | 実行内容による | △（base64 で約 1.3 倍に膨らむ） | 不要（CDP のために構築済み） | 可 | 軽量データ向け |
| `Runtime.addBinding` | JS→VBA | なし（返事は `Runtime.evaluate` / `Runtime.callFunctionOn`） | △（base64 で約 1.3 倍に膨らむ） | 不要（CDP のために構築済み） | 可 | 軽量データ向け。基本は `Runtime.evaluate` / `Runtime.callFunctionOn` とセット |
| SharedBuffer | 双方向 | なし | ◎ | 不要 | 不可 | **大容量の本命** |
| `Browser.downloadProgress`（blob URL を DL） | JS→ディスク | ― | ○（ディスク保存のコストあり） | 不要 | **可** | スクレイピングなら大抵これで十分 |
| `AddHostObjectToScript` | JS→VBA（RPC） | **あり** | ×（文字列なら可） | 不要 | 不可 | 返事が必要な合図に有力 |

「通常の Edge」が「不可」のものは、WebView2（UserForm に埋め込んだ WebView2）でだけ使えます。

なお、WebView2 にはもう 1 つ、`WebMessage`（ページ側は `chrome.webview.postMessage`）による双方向のやり取りもあります。Microsoft が推奨する方法ですが、受け取りのために追加のサンクの儀式が要ります。CDP を構築済みのこのツールにとっては、上の選択肢と大差がないため、取り込んでいません。

## 各選択肢の特徴

### `Runtime.evaluate` / `Runtime.callFunctionOn`

VBA から JavaScript を実行し、その結果を受け取ります。このツールでは [`jsEval`](/guides/javascript) がこれにあたります（`CDPContext.jsEval` は主に `Runtime.evaluate`、`CDPElement.jsEval` は `Runtime.callFunctionOn`）。WebDriver BiDi の `jsEval` も、`script.evaluate` / `script.callFunction` で同じことをします。

```vb
Debug.Print t.jsEval("document.title")
Debug.Print t.jsEval("fetch('/api/items').then(r => r.text())", awaitPromise:=True)
```

- **メリット**
  - 追加の準備が要りません。通常の Edge / Chrome でも、WebView2 でも使えます
  - 戻り値を受け取れます。`awaitPromise` を付ければ、Promise が終わるまで待ってから受け取れます
- **デメリット**
  - VBA から呼びかける一方通行です。ページ側で何かが起きたことを知るには、VBA から何度も聞きに行く必要があります
  - 結果は JSON の文字列で届きます。バイナリは base64 にするので、約 1.3 倍に膨らみます。大きなデータには向きません

値を取り出す、ボタンを押す、といった普段の操作は、ほとんどこれで足ります。

### `Runtime.addBinding`

名前を指定して `Runtime.addBinding` を送ると、ページに同じ名前のグローバル関数ができます。ページの JavaScript がその関数を文字列付きで呼ぶと、VBA に `Runtime.bindingCalled` イベント（関数名と文字列）が届きます。

```vb
Dim params As New Dictionary
params.Add "name", "notifyVBA"
t.ExecuteCDP "Runtime.addBinding", params

' ページ側からは、次のように呼ぶ
' notifyVBA(JSON.stringify({ status: "done", count: 42 }))
```

`Runtime.bindingCalled` は、[イベント購読](/guides/events)（`BrowserEvents` や、拡張クラスの `ExEvent`）で受け取ります。このツール自身も、WebDriver BiDi の結果をページから受け取るのにこの方法を使っています（[BiDi 登場秘話](/stories/DevHistory/bidi-story)）。

- **メリット**
  - ページ側から VBA へ、好きなときに合図を送れます。VBA から何度も聞きに行く必要がありません
  - 通常の Edge / Chrome でも使えます
- **デメリット**
  - 戻り値がありません。返事をするには、VBA から `jsEval` で改めて呼びかけます
  - 渡せるのは文字列 1 つだけです。バイナリは base64 にするので、大きなデータには向きません

合図は `Runtime.addBinding`、返事は `jsEval`、と組み合わせて使うのが基本です。

`Runtime.addBinding` が普通のブラウザで使えるようになった経緯は、[技術資料: Web とローカルの双方向やり取り](/stories/DebugRouteTech/Two-wayInteractionStory)にまとめています。

### SharedBuffer

WebView2 の `ICoreWebView2SharedBuffer` で、VBA とページの JavaScript が同じメモリ領域を読み書きします。各メンバーは [WebView2モードでできること](/webview2/capabilities) の「共有バッファ」を参照してください。

次のデモは、ページの HTML を JavaScript が共有バッファへ書き込み、VBA が読み出します。JavaScript 側には「書き終わった」ことを知らせる仕組みが無いので、書いたバイト数を `window.__sbRes` に置いてもらい、VBA は `jsEval` でそれを待ちます。

```vb
Sub SharedBufferDemo()
    Dim wv2 As CDPCoreViaWebView2, t As CDPContext
    WebView2Form.StartCDPModeWebView2
    Set wv2 = WebView2Form.ThisWebView2
    Set t = WebView2Form.ThisCDPContext
    t.navigate "https://example.com"

    '1. 受け取る側（JS）のリスナーを先に登録する
    '   HTML を UTF-8 で書き込み、releaseBuffer してから、書いたバイト数を window.__sbRes に置く
    t.jsEval "chrome.webview.addEventListener('sharedbufferreceived', e => {" & _
             " const buf = e.getBuffer();" & _
             " const n = new TextEncoder().encodeInto(document.documentElement.outerHTML, new Uint8Array(buf)).written;" & _
             " chrome.webview.releaseBuffer(buf);" & _
             " window.__sbRes = n; }); 0"

    '2. 共有バッファ（1MB）を作り、ページへ渡す
    Dim sb As WebView2SharedBuffer
    Set sb = wv2.CreateSharedBuffer(CLngLng(1024& * 1024&))
    If sb Is Nothing Then Exit Sub
    t.jsEval "window.__sbRes = undefined; 0"
    sb.PostToScript SharedBuffer_ReadWrite

    '3. JS が書き終わるのを待つ（書いたバイト数が返る）
    Dim written As Long
    written = Val(WaitJs(t, "window.__sbRes", 10))

    '4. 先頭から読み出して、UTF-8 から文字列に戻す
    Dim bytes() As Byte
    bytes = sb.ReadBytes(written)
    Dim conv As New CharacterCodeConversion
    Debug.Print conv.BytesToString(bytes)

    Set sb = Nothing
    Unload WebView2Form
End Sub

'JS の変数に値が入るまで待つ（タイムアウト時は空文字）
Private Function WaitJs(t As CDPContext, ByVal jsVar As String, ByVal timeoutSec As Double) As String
    Dim t0 As Single: t0 = Timer
    Dim v As Variant
    Do
        v = t.jsEval(jsVar & " === undefined ? '' : String(" & jsVar & ")")
        If Not IsError(v) Then
            If Len(v) > 0 Then WaitJs = v: Exit Function
        End If
        DoEvents
    Loop Until Timer - t0 > timeoutSec
End Function
```

VBA → JS の向きも同じです。VBA が `WriteBytes` で書いてから `PostToScript` し、JavaScript が `TextDecoder` などで読みます。

::: tip 大きなデータはファイルへ直接書く
数百 MB のデータを `ReadBytes` で VBA の Byte 配列に丸ごと読むと、メモリを圧迫します。ファイルに保存したいだけなら、`get_Buffer`（共有メモリの先頭アドレス）を Windows API の `WriteFile` にそのまま渡せば、VBA の配列を経由せずに書き出せます。
:::

動作確認用のテストは、dev-xlflow ブランチの `ForDevelopers/OperationCheck/WebView2/WebView2SharedBuffer/` にあります。`Test_WebView2SharedBuffer.bas` は読み書き・往復・100MB・サイズ上限などの単体テスト、`Test_WV2SharedBufferHLS.bas` は HLS の動画（映像と音声）を JavaScript で集め、共有バッファ経由でファイルに保存するテストです。

- **メリット**
  - 大きなデータを、base64 にもせず、ディスクも経由せずに渡せます
  - VBA からもページからも、読み書きできます
- **デメリット**
  - WebView2 でだけ使えます。それも、トップレベルのページ（UserForm に表示しているページ）だけです
  - 「書き込み終わり」の合図がありません。`Runtime.addBinding` で合図を送るか、同期の `jsEval` で書き込みを済ませてから読みます
  - 確保したサイズは、あとから変えられません。上限は 2,147,418,112 バイト（2GB − 64KB）です
    - 64KB は、Windows の割り当ての粒度です。Chromium が共有メモリのサイズをこの単位に切り上げてから上限と比べるため、2GB ちょうどより 64KB 少なくなります。64KB になった経緯は、Raymond Chen の [Why is address space allocation granularity 64KB?](https://devblogs.microsoft.com/oldnewthing/20031008-00/?p=42223)（The Old New Thing、2003年10月8日）で説明されています

大きなデータを VBA の中だけで受け渡したいときの本命です。

### `Browser.downloadProgress`（blob URL をダウンロード）

ページの JavaScript でデータを `Blob` にし、`URL.createObjectURL` で作った `blob:` の URL をダウンロードさせます。VBA 側は、`Browser.setDownloadBehavior`（`eventsEnabled` を `true`）でダウンロードのイベントを有効にし、`Browser.downloadProgress` で完了を待ってから、保存されたファイルを読みます。

- **メリット**
  - 通常の Edge / Chrome でも使えます
  - 大きなデータも渡せます。CSV などのファイルとして欲しいなら、そのまま成果物になります
- **デメリット**
  - 一度ディスクに保存するので、その分の時間がかかります。VBA で使うには、ファイルを読み直す必要もあります
  - ダウンロード先や、ダウンロードの設定を管理する必要があります

スクレイピングで「データをファイルとして持ち帰る」なら、大抵これで十分です。ダウンロードの監視には、Extensions ブランチの `exCDP_DownloadWatcher` も使えます。

### `AddHostObjectToScript`

WebView2 の `AddHostObjectToScript` で、VBA のクラスのオブジェクトをページに公開します。ページの JavaScript から `chrome.webview.hostObjects.{名前}.{メソッド名}(...)` で呼ぶと、VBA のメソッドが動き、その戻り値が JavaScript に返ります。IE 時代の `window.external` や、`onclick` に VBA のオブジェクトを代入するやり方に近いものです。

次のデモは、UserForm に表示したページのボタンから VBA を呼び、Excel の A1 セルに値を書き込んで、A1:A10 の合計をページに表示します。

```vb
' クラスモジュール「clsWebBridge」
' JS から呼ばれるメソッドの引数は ByVal にする（ByRef のままだと型の不一致エラーになりやすい）
Public Function UpdateSheet(ByVal targetCell As String, ByVal numValue As Long) As Double
    ActiveSheet.Range(targetCell).Value = numValue
    UpdateSheet = Application.WorksheetFunction.Sum(ActiveSheet.Range("A1:A10"))
End Function
```

```vb
' 標準モジュール
Sub HostObjectDemo()
    With WebView2Form
        '1. WebView2 を起動する
        .StartCDPModeWebView2

        '2. クラスのオブジェクトを "vba" という名前でページに公開する
        Dim bridge As New clsWebBridge
        .ThisWebView2.AddHostObjectToScript "vba", bridge

        '3. ボタンが 1 つだけのページを表示する
        Dim html As String
        html = "<button onclick=""send()"">A1 に 50 を書き込む</button><div id=""out""></div>" & _
               "<script>async function send() {" & _
               " const sum = await chrome.webview.hostObjects.vba.UpdateSheet('A1', 50);" & _
               " document.getElementById('out').innerText = 'A1:A10 の合計 = ' + sum; }</script>"

        Dim params As New Dictionary
        params.Add "frameId", .ThisCDPContext.CurrentTargetID
        params.Add "html", html
        .ThisCDPContext.ExecuteCDP "Page.setDocumentContent", params

        '4. フォームを開いたまま、VBA の処理を終える
        .show vbModeless
    End With
End Sub
```

ページの JavaScript からの呼び出しは、VBA がウィンドウメッセージを処理しているときに実行されます。このデモのように、モードレスでフォームを開いて VBA の処理を終えておけば、ボタンを押すたびに `UpdateSheet` が呼ばれます。VBA のマクロの実行中に呼ばれた場合は、`DoEvents` などでメッセージが処理されるまで待たされます。

動作確認用のテストは、dev-xlflow ブランチの `ForDevelopers/OperationCheck/WebView2/WebView2HostObject/` にあります。`Excel連携/` はこのデモの元になった UserForm のデモ（フォルダ選択ダイアログや音声読み上げも VBA に頼む例）、`HeadlessTest/` は Property や Public 変数の読み書き、同期で呼ぶ `hostObjects.sync`、`RemoveHostObjectFromScript` などのテストです。

- **メリット**
  - ページから VBA を呼び、戻り値を受け取れます
  - VBA から何度も聞きに行く必要がありません
- **デメリット**
  - WebView2 でだけ使えます
  - VBA のメソッドが終わるまで、WebView2 は止まったままです（VBA も WebView2 の UI も、同じ 1 つのスレッドで動くため）。重い処理には向きません
  - 引数や戻り値は、文字列などの小さな値向けです。大きなデータには向きません

ページ上のボタンから VBA の処理を呼び、その結果をページに表示したい、といった「返事が要る合図」に向いています。

## 関連

- [JavaScript 実行](/guides/javascript)
- [イベント購読](/guides/events)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [WebView2モードでできること](/webview2/capabilities)

## 参考資料

### IE 時代

* [Handling Events in Visual Basic Applications](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752045(v=vs.85))（既定のメソッドを持つクラスを、要素の `onclick` に代入する例。イベントは WebBrowser コントロールまでバブリングしない）
* [How to sink HTML document events for WebBrowser host（KB246247 のアーカイブ）](https://mskb.pkisolutions.com/kb/246247)（`DocumentComplete` のあとに `WithEvents` の `HTMLDocument` を受け取り、`doc_onclick` で受ける例）
* [DWebBrowserEvents2 interface](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa768283(v=vs.85))（`BeforeNavigate2`、`DocumentComplete`、`NewWindow2`、`OnQuit` などのイベント一覧）
* [IHTMLWindow2::execScript method](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741364(v=vs.85))（IE11 からはサポートされず、`eval` を使う）
* [Compatibility changes in IE11](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/dev-guides/bg182625(v=vs.85))（削除された API として `window.execScript`、代わりは `eval`）
* [IHTMLDocument::Script property](https://learn.microsoft.com/en-us/previous-versions/aa752642(v=vs.85))（スクリプトエンジンへのハンドルを `IDispatch` で受け取る）
* [IDocHostUIHandler::GetExternal method](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa753256(v=vs.85))（ホストの `IDispatch` を MSHTML に渡す）
* [WebBrowser.ObjectForScripting Property](https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.webbrowser.objectforscripting)（.NET の WebBrowser で、`window.external` からホストのオブジェクトを呼ぶ）

### CDP / WebView2

* [Runtime.addBinding](https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#method-addBinding)（呼ぶと `Runtime.bindingCalled` が届くグローバル関数を追加する）
* [Browser.setDownloadBehavior](https://chromedevtools.github.io/devtools-protocol/tot/Browser/#method-setDownloadBehavior) / [Browser.downloadProgress](https://chromedevtools.github.io/devtools-protocol/tot/Browser/#event-downloadProgress)
* [ICoreWebView2::AddHostObjectToScript](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2#addhostobjecttoscript)
* [ICoreWebView2Environment12::CreateSharedBuffer](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/reference/win32/icorewebview2environment12#createsharedbuffer)
* [platform_shared_memory_region_win.cc](https://chromium.googlesource.com/chromium/src/+/565c30fae47aba4a1e7d9bd986ae9fe2a5ac7445/base/memory/platform_shared_memory_region_win.cc#291)（Chromium。共有メモリのサイズを 64KB 単位に切り上げてから、2^31 − 1 以下かを確かめる）
* [Why is address space allocation granularity 64KB?](https://devblogs.microsoft.com/oldnewthing/20031008-00/?p=42223)（The Old New Thing、2003年10月8日。Windows の割り当ての粒度が 64KB になった経緯）
* [ネイティブ コードと Web コードの相互運用](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/how-to/communicate-btwn-web-native)（`PostWebMessage` / `chrome.webview.postMessage` によるやり取り）
