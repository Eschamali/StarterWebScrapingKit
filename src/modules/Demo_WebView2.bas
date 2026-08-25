Attribute VB_Name = "Demo_WebView2"
'***************************************************************************************************
'                   WebView2 経由でのCDP制御デモです
'
'   基本的には以下の流れで利用可能です
'   1. Demo_WebView2Init                          … WebView2(Environment/Controller/View)を初期化
'   2. Demo_WebView2SubscribeEvents(任意)          … 個別にCDPイベント名を購読
'   3. Demo_WebView2CDP_RuntimeEvaluate            … reattach経由でCDPコマンドを一往復
'
'   ★注意事項★
'   ・WebView2はCDP-over-WebSocketと異なり、`GetDevToolsProtocolEventReceiver`が
'     「イベント名ごとの個別登録」のみをサポートします。ドメインをenableするだけでは
'     イベントは流れてきません。必要なイベント名ごとに`SubscribeCdpEvent`を呼んでください。
'   ・このパスで受け取るイベントには`sessionId`が付与されないため、`CDPBrowserEvent`側で
'     発火します(`CDPContextEvent`側では発火しません)。
'   ・コールバック待ち中(コマンド送信?完了、イベント購読中)にVBEでブレーク/ステップ実行すると
'     Excelがクラッシュする可能性があります。検証中はブレークしないよう注意してください。
'***************************************************************************************************
Option Explicit
Option Private Module

'再利用のため、モジュール変数として保持
Public g_webview2Obj As CDPCoreViaWebView2

'***************************************************************************************************
'* 機能　　：WebView2(Environment/Controller/View)を初期化します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`WebView2Loader.dll`はこのプロジェクト独自には同梱しません。Excelの
'            Power Query統合アドインに同梱されている実物を実行時に探索して使用します。
'            見つからない場合は失敗します(Power Query for Excelの導入状況を確認してください)。
'***************************************************************************************************
Sub Demo_WebView2Init()
    Set g_webview2Obj = New CDPCoreViaWebView2

    If Not g_webview2Obj.ConnectCDP(ShSetting01_StartBrowser.CurrentUserName) Then
        Debug.Print "WebView2の初期化に失敗しました。WebView2Loader.dllが見つからない、" & _
                    "またはEnvironment/Controllerの生成に失敗した可能性があります。"
        Set g_webview2Obj = Nothing
        Exit Sub
    End If

    Debug.Print "WebView2 の初期化が完了しました。"
End Sub

'***************************************************************************************************
'* 機能　　：任意のCDPイベント名を個別に購読します(必要な分だけ、複数回呼び出してください)
'***************************************************************************************************
Sub Demo_WebView2SubscribeEvents()
    If g_webview2Obj Is Nothing Then
        Debug.Print "先に Demo_WebView2Init を実行してください。"
        Exit Sub
    End If

    Debug.Print "Page.loadEventFired の購読: " & g_webview2Obj.SubscribeCdpEvent("Page.loadEventFired")
    Debug.Print "Network.requestWillBeSent の購読: " & g_webview2Obj.SubscribeCdpEvent("Network.requestWillBeSent")
    Debug.Print "Network.loadingFinished の購読: " & g_webview2Obj.SubscribeCdpEvent("Network.loadingFinished")
    Debug.Print "Target.attachedToTarget の購読: " & g_webview2Obj.SubscribeCdpEvent("Target.attachedToTarget")
End Sub

'***************************************************************************************************
'* 機能　　：`CDPBrowser.reattach`経由でWebView2をCDP制御し、`Runtime.evaluate`を一往復させます
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebView2は「browser-endpoint配下の複数タブ」という構造を持たないため、`getTab`/
'            `CDPContext`は使わず、`CDPBrowser`(browser-level/セッションレス)のみで完結させます
'***************************************************************************************************
Sub Demo_WebView2CDP_RuntimeEvaluate()
    If g_webview2Obj Is Nothing Then
        Debug.Print "先に Demo_WebView2Init を実行してください。"
        Exit Sub
    End If

    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    Dim c As New CDPBrowser
    If Not c.reattach(UserName, WebView2Mode:=g_webview2Obj) Then
        MsgBox "WebView2への reattach に失敗しました。", vbCritical, "Chrome DevTools Protocol"
        Exit Sub
    End If

    '1. ページ遷移
    Dim navParams As New Dictionary
    navParams.Add "url", "https://example.com/"
    c.ExecuteCDP "Page.navigate", navParams

    '2. document.title を取得
    Dim evalParams As New Dictionary
    evalParams.Add "expression", "document.title"

    Dim resultNode As BiDiCDPJson
    Set resultNode = c.ExecuteCDP("Runtime.evaluate", evalParams)

    Debug.Print "document.title = " & resultNode.NodeKey("result").StringKey("value")
End Sub

'***************************************************************************************************
'* 機能　　：`CallDevToolsProtocolMethodForSession`経路(sessionId指定)の確認手順メモです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：`sessionId`を得るには、まず`Target.setAutoAttach`を有効化し、クロスオリジンiframe等で
'            OOPIF(別プロセスのフレーム)が発生した際に届く`Target.attachedToTarget`イベントの
'            JSONから`sessionId`を回収する必要があります。手動検証用の骨組みのみ用意しています。
'***************************************************************************************************
Sub Demo_WebView2CDP_ForSession_Note()
    If g_webview2Obj Is Nothing Then
        Debug.Print "先に Demo_WebView2Init を実行してください。"
        Exit Sub
    End If

    Debug.Print "1. Demo_WebView2SubscribeEvents で Target.attachedToTarget を購読してください"
    Debug.Print "2. CDPBrowser.ExecuteCDP ""Target.setAutoAttach"" で" & _
                " {""autoAttach"":true,""waitForDebuggerOnStart"":false,""flatten"":true} を送ってください"
    Debug.Print "3. クロスオリジンiframeを含むページへ Page.navigate してください"
    Debug.Print "4. Immediate Windowで CDPContextEvent/CDPBrowserEvent 経由で受け取った" & _
                " Target.attachedToTarget の生JSONから sessionId を確認してください"
    Debug.Print "5. その sessionId を含む CDP-Jsonコマンド文字列を、CDPCoreViaWebView2." & _
                "SendCommandCDP に直接渡すと、CallDevToolsProtocolMethodForSession経路(vtableスロット99)を通ります"
End Sub

Sub ネットワークイベントの確認ForWebView2()
    '必要な変換オブジェクトを用意
    Dim CharConvObj As New CharacterCodeConversion

    '設定シートに基づくブラウザ立ち上げ
    Dim Demo_NetworkEvent As CDPContext
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    Dim c As New CDPBrowser
    If Not c.reattach(UserName, WebView2Mode:=g_webview2Obj) Then
        MsgBox "WebView2への reattach に失敗しました。", vbCritical, "Chrome DevTools Protocol"
        Exit Sub
    End If
    Set Demo_NetworkEvent = c.newTab(setMain:=True)

    '一部の非同期イベントのみキャプチャするようにフィルターを設定
    '※未設定の場合は、全キャプチャとなります。このDemoの場合は、下記2つをコメントアウトすると、全キャプチャとなります
    Demo_NetworkEvent.SetFilterEvents = "Network.requestWillBeSent"
    Demo_NetworkEvent.SetFilterEvents = "Network.loadingFinished"


    '-------------------------------- 機能1：イベントキャプチャを有効化する --------------------------------
    Set Demo_NetworkEvent.BrowserEvents = New Dictionary        '`New Dictionary`を渡すことで、新規イベントキャプチャが可能になる。


    'ネットワークイベント受信を有効化する
    Demo_NetworkEvent.ExecuteCDP "Network.enable"

    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/excel/vba/file/file11.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.InheritanceCDPBrowser.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "Event.json"


    '-------------------------------- 機能2：セーブデータを作成し、イベントキャプチャを無効化する --------------------------------
    Dim SaveDataEvents As Dictionary: Set SaveDataEvents = Demo_NetworkEvent.BrowserEvents  'セーブデータ作成
    Set Demo_NetworkEvent.BrowserEvents = Nothing               '`Nothing`を渡すことで、イベントを破棄するようになる


    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/youtube/20200714b.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.InheritanceCDPBrowser.TakeEvents

    'イベント情報をDownloadsフォルダに保存しますが、無効中なので0バイトになります
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "NotEvent.json"


    '-------------------------------- 機能3：セーブデータを読み込み、そこからイベントキャプチャを再開する --------------------------------
    Set Demo_NetworkEvent.BrowserEvents = SaveDataEvents        '既存のセーブデータを読み込む


    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/index.stm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.InheritanceCDPBrowser.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "EventFromSaveData.json"


    'ブラウザを閉じる。demo終了
    Demo_NetworkEvent.InheritanceCDPBrowser.quit
End Sub

