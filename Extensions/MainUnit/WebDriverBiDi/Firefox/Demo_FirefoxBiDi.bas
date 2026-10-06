Attribute VB_Name = "Demo_FirefoxBiDi"
'==============================================================================================================
'   Firefoxの直WebDriverBiDi制御Demoです。
'   事前に、
'   "firefox.exe" -no-remote --remote-debugging-port=0 -profile "%UserProfile%\AppData\Roaming\Mozilla\Firefox\Profiles\FirefoxDebug"
'   等でデバッグブラウザを起動しておく必要があります
'==============================================================================================================
Option Explicit
Option Private Module



'***************************************************************************************************
'                                   ■■■ 接続アシスト ■■■
'***************************************************************************************************
'* 機能　　：デバッグFirefoxへWebSocket接続を行います
'---------------------------------------------------------------------------------------------------
'* 引数　　：ReusingWinSockHandle   `True`で、テーブルに記録されてる値で、Winsockハンドルを使いまわします
'* 返り値　：接続済みデバッグFirefoxのWebSocketオブジェクト
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebSocket切断 or `session.end`してない場合は、第一引数にて、`True`が可能です
'***************************************************************************************************
Private Function ConnectFirefox(Optional ReusingWinSockHandle As Boolean) As CDPCoreViaWebSocket
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. 指定のWebSocketForBiDiへ接続
    Set ConnectFirefox = New CDPCoreViaWebSocket
    ConnectFirefox.DirectWebDriverBiDiMode = True

    '3. 接続方法に応じて分岐
    If ReusingWinSockHandle Then
        '有効な、WinSockハンドルを使いまわす場合用
        ConnectFirefox.ReConnectCDP UserName, True
    Else
        '"%UserProfile%\AppData\Roaming\Mozilla\Firefox\Profiles\FirefoxDebug\WebDriverBiDiServer.json"から特定してください
        Debug.Print ConnectFirefox.ConnectCDP(UserName, "/session", 57221)
    End If
End Function



'***************************************************************************************************
'                               ■■■ Demoプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：イベントキャプチャに関するDemoコード(BiDi版)です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：CDP版の「ネットワークイベントの確認」をBiDiの`network`ドメインを用いて再現したデモです。
'*           `session.subscribe` で `network` 関連イベントを購読し、結果をJSON出力します。
'***************************************************************************************************
Sub checkNetworkEvents()
    '必要な変換オブジェクトを用意
    Dim CharConvObj As New CharacterCodeConversion

    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. 繋げたWebSocketオブジェクトを`reattachWebSocket`メソッドに渡す
    Dim WebSocketFirefox As New WebDriverBiDiMode
    WebSocketFirefox.reattach UserName, , ConnectFirefox

    '3. 既存タブに接続
    Dim Demo_NetworkEvent As WebDriverBiDiContext
    Set Demo_NetworkEvent = WebSocketFirefox.getTab(setMain:=True)


    '-------------------------------- 機能1：イベントキャプチャを有効化する --------------------------------
    '`New Dictionary`を渡すことで、内部で非同期イベントの蓄積を開始する
    Dim resultBiDi As Dictionary
    Set Demo_NetworkEvent.ThisWebDriverBiDiMode.BiDiEvents = New Dictionary

    'BiDi側でネットワークイベントを購読開始する
    Dim paramsBiDi As Dictionary
    Set paramsBiDi = New Dictionary
    Demo_NetworkEvent.SubscribeBiDiEvent = Array("network.beforeRequestSent", "network.responseCompleted", "log.entryAdded")

    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/excel/vba/file/file11.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う (念のため待機後にも余波を回収)
    Demo_NetworkEvent.ThisWebDriverBiDiMode.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.ThisWebDriverBiDiMode.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_Event.json"


    '-------------------------------- 機能2：セーブデータを作成し、イベントキャプチャを無効化する --------------------------------
    Dim SaveDataEvents As Dictionary: Set SaveDataEvents = Demo_NetworkEvent.ThisWebDriverBiDiMode.BiDiEvents  'セーブデータ作成
    Set Demo_NetworkEvent.ThisWebDriverBiDiMode.BiDiEvents = Nothing               '`Nothing`を渡すことで、イベント記録状態を破棄する


    'URL遷移
    Demo_NetworkEvent.navigate "http://officetanaka.net/youtube/20200714b.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.ThisWebDriverBiDiMode.TakeEvents

    'イベント情報をDownloadsフォルダに保存しますが、無効中なので破棄状態（0バイト等）になります
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.ThisWebDriverBiDiMode.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_NotEvent.json"


    '-------------------------------- 機能3：セーブデータを読み込み、そこからイベントキャプチャを再開する --------------------------------
    Set Demo_NetworkEvent.ThisWebDriverBiDiMode.BiDiEvents = SaveDataEvents        '既存のセーブデータを読み込む

    'URL遷移
    Demo_NetworkEvent.navigate "http://officetanaka.net/index.stm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.ThisWebDriverBiDiMode.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.ThisWebDriverBiDiMode.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_EventFromSaveData.json"


    'ブラウザを閉じる。demo終了
    Demo_NetworkEvent.ThisWebDriverBiDiMode.quit
End Sub



'***************************************************************************************************
'                               ■■■ リアタッチDemo ■■■
'***************************************************************************************************
'* 機能　　：複数プロシージャをまたがった段階的な処理を行う際の再接続Demoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：単一プロシージャで完結出来ない場面がきっとあるはずです。途中でセキュリティ認証による手作業が入ったりなど...
'            そういった場面でも、デバックブラウザで起動済みへ再接続するDemoです
'---------------------------------------------------------------------------------------------------
'* 注意事項：事前にデバッグFirefoxの起動が必要です
'***************************************************************************************************
Sub demoReattachmentPart1()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. 繋げたWebSocketオブジェクトを`reattachWebSocket`メソッドに渡す
    Dim WebSocketFirefox As New WebDriverBiDiMode
    WebSocketFirefox.reattach UserName, , ConnectFirefox

    '3. 既存タブに接続
    Dim First As WebDriverBiDiContext
    Set First = WebSocketFirefox.getTab(setMain:=True)

    '4. GoogleTopページへ遷移
    First.navigate "https://developer.mozilla.org/en-US/docs/Web/WebDriver/How_to/Create_BiDi_connection#launching_the_browser"

    '5. セッションのみ切断
    WebSocketFirefox.sessionEnd
End Sub

'***************************************************************************************************
'* 機能　　：WebDriverBiDi制御用タブの接続まで担うリアタッチです
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、WebDriverBiDi制御用のタブ接続までです。その後のContext(タブ)接続は、手動で`getTab` OR `newTab`で出来ます
'            ・ブラウザのパイプハンドルが生きてない場合は、エラーになります。`demoReattachmentPart1`からやり直しです
'            ・WebDriverBiDi制御用タブが無くなっても、`WebDriverBiDiMode`からの`reattach`で、再始動が可能です
'***************************************************************************************************
Sub demoReattachmentPart2()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. 繋げたWebSocketオブジェクトを`reattachWebSocket`メソッドに渡す
    Dim WebSocketFirefox As New WebDriverBiDiMode
    WebSocketFirefox.reattach UserName, , ConnectFirefox

    '3. 未接続のタブに接続
    '※この時、必ず`setMain:=True`とすること。必要に応じて検索条件(URLマッチ等)も設定して下さい
    Dim ReattachmentTab As WebDriverBiDiContext
    Set ReattachmentTab = WebSocketFirefox.getTab(setMain:=True)
'    Set ReattachmentTab = WebSocketFirefox.newTab(setMain:=True)   '新しいタブ生成からでもOK

    '4．別ページに遷移して終了
    ReattachmentTab.navigate "https://kemono-friends-20170110.jp/"
End Sub

'***************************************************************************************************
'* 機能　　：最後にWebDriverBiDiで制御したタブの接続まで担うリアタッチです
'---------------------------------------------------------------------------------------------------
'* 注意事項：最後にWebDriverBiDiで制御したタブが失ってる場合は失敗します
'***************************************************************************************************
Sub demoReattachmentPart2ForTab()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. リアタッチとして起動
    Dim Reattachment As New WebDriverBiDiContext
    If Not Reattachment.reattach(UserName, , ConnectFirefox(True)) Then MsgBox "「" & UserName & "」に接続できませんでした。`BiDi-context`情報がお亡くなりです。", vbCritical, "WebDriver BiDi": Exit Sub

    '3. 別ページに遷移
    Reattachment.navigate "https://w3c.github.io/webdriver-bidi/"

    '4. 後始末
    Reattachment.ThisWebDriverBiDiMode.sessionEnd
End Sub
