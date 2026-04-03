Attribute VB_Name = "Demo_WebView2"
Option Explicit
'***
' Demo_WebView2.bas
' WebView2Core.cls の動作確認デモ
'
' 【実行前の準備】
' 1. VBEプロジェクトに WebView2Core.cls と WebView2Callbacks.bas をインポート
' 2. UserForm に Frame または PictureBox を配置し、そのhWndを使うか、
'    下記のように UserForm 自体の hWnd を親ウィンドウとして指定する
' 3. ExcelウィンドウのhWndを使う最小デモ → Sub TestWebView2Simple を実行
'***



'***************************************************************************************************
'                               ■■■ 設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：設定シートから、パラメーターを読み込んで、WebView2を起動するヘルパープロシージャです
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebView2Browser
'* 引数　　：StartURL                       ブラウザ起動時にアクセスしたいURL。指定しない場合は、空ページ(abount:blank)になります。
'            SwtchUser                      マルチインスタンス用に別ユーザーを指定するときに使用します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBEによるハードコーディングではなく、設定シートから読み込む方式により、ユーザー側からも手軽に設定変更ができます
'* 注意事項：Demoモジュールにあるコードですが、他の部分で共用してるため、消さずにどこかにカット&ペーストしておくとよいでしょう
'***************************************************************************************************
Public Function 設定シートからのWebView2起動(Optional StartURL As String, Optional SwitchUser As String) As WebView2Browser
    #If ViewLog = 1 Then
        WV2logView = True
    #Else
        WV2logView = False
    #End If
    

    '設定シートの各セルから設定値を取得し、適用
    With ShSetting01_StartBrowser
        '第2引数が省略ならシート側の設定を適用
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .Range(.UseRangeName(2, "Demo_WebView2.設定シートからのWebView2起動")).value, SwitchUser)

        'ブラウザ起動
        Set 設定シートからのWebView2起動 = New WebView2Browser
        設定シートからのWebView2起動.start UseDataDir, .Range(.UseRangeName(12, "Demo_WebView2.設定シートからのWebView2起動")).value, StartURL, .Range(.UseRangeName(3, "Demo_WebView2.設定シートからのWebView2起動")).value
    End With
End Function

Sub WebView2による冒険の始まり()
    '設定シートに基づくブラウザ立ち上げ
    Dim HelloWorldAutomationBrowser As WebView2Browser: Set HelloWorldAutomationBrowser = 設定シートからのWebView2起動

    '↓ここから、あなたのイメージをコードに落とし込む↓




    'ブラウザを正常に閉じる
    Unload WebView2InExcelForm
    HelloWorldAutomationBrowser.quit
End Sub



'***************************************************************************************************
'                               ■■■ Demoプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：イベントキャプチャに関するDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：例えば、認証用URLのNetwork.loadingFinished を検知したら、そこの requestId から `Network.getResponseBody` を実行しToken入手なんてことが可能です。(でも、Token抽出とかはNetwork.getCookies や DOMStorage.getDOMStorageItems 等が楽です。)
'* 注意事項：・ここでは、ネットワークイベントのデモですが、他のイベントも同じ操作でとらえることができます
'            ・WebView2の仕様上、イベント購読設定と対応する`○○.enable`の設定が必要になります
'***************************************************************************************************
Sub ネットワークイベントの確認()
    '必要な変換オブジェクトを用意
    Dim JsonDicObj As New WebJsonConverter
    Dim CharConvObj As New CharacterCodeConversion:
    
    '設定シートに基づくブラウザ立ち上げ
    Dim Demo_NetworkEvent As WebView2Browser: Set Demo_NetworkEvent = 設定シートからのWebView2起動

    'ネットワーク関連のイベントをとりあえず全部購読する設定を施す
    'https://chromedevtools.github.io/devtools-protocol/tot/Network/
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.dataReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.eventSourceMessageReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.loadingFailed"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.loadingFinished"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.requestServedFromCache"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.requestWillBeSent"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.responseReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webSocketClosed"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webSocketCreated"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webSocketFrameError"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webSocketFrameReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webSocketFrameSent"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webSocketHandshakeResponseReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webSocketWillSendHandshakeRequest"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webTransportClosed"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webTransportConnectionEstablished"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.webTransportCreated"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.deviceBoundSessionEventOccurred"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.deviceBoundSessionsAdded"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directTCPSocketAborted"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directTCPSocketChunkReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directTCPSocketChunkSent"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directTCPSocketClosed"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directTCPSocketCreated"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directTCPSocketOpened"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketAborted"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketChunkReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketChunkSent"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketClosed"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketCreated"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketJoinedMulticastGroup"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketLeftMulticastGroup"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.directUDPSocketOpened"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.policyUpdated"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.reportingApiEndpointsChangedForOrigin"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.reportingApiReportAdded"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.reportingApiReportUpdated"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.requestWillBeSentExtraInfo"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.resourceChangedPriority"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.responseReceivedEarlyHints"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.responseReceivedExtraInfo"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.signedExchangeReceived"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.trustTokenOperationDone"
    Demo_NetworkEvent.SubscribeDevToolsProtocolEvent "Network.requestInterceptedDeprecated"


    '-------------------------------- 機能1：イベントキャプチャを有効化する --------------------------------
    Set Demo_NetworkEvent.BrowserEvents = New Dictionary        '`New Dictionary`を渡すことで、新規イベントキャプチャが可能になる。

    
    'ネットワークイベント受信を有効化する
    Dim ResultCDP As Dictionary: Set ResultCDP = Demo_NetworkEvent.invokeMethod("Network.enable")
    
    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/excel/vba/file/file11.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "WebView2Event.json"


    '-------------------------------- 機能2：セーブデータを作成し、イベントキャプチャを無効化する --------------------------------
    Dim SaveDataEvents As Dictionary: Set SaveDataEvents = Demo_NetworkEvent.BrowserEvents  'セーブデータ作成
    Set Demo_NetworkEvent.BrowserEvents = Nothing               '`Nothing`を渡すことで、イベントを破棄するようになる(※ただし、イベント購読自体は継続)


    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/youtube/20200714b.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存しますが、無効中なので0バイトになります
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "WebView2NotEvent.json"


    '-------------------------------- 機能3：セーブデータを読み込み、そこからイベントキャプチャを再開する --------------------------------
    Set Demo_NetworkEvent.BrowserEvents = SaveDataEvents        '既存のセーブデータを読み込む

    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/index.stm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "WebView2EventFromSaveData.json"
    
    '全部のイベント購読を止める
    Demo_NetworkEvent.UnsubscribeDevToolsProtocolEvent


    'ブラウザを閉じる。demo終了
    Unload WebView2InExcelForm
    Demo_NetworkEvent.quit
End Sub

'***************************************************************************************************
'* 機能　　：日本語に関するDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：id属性やname属性に日本語が使われてるサイトでの動作テストです。コードは、`https://qiita.com/yaju/items/0807cc762af4a0568806`を参考にしてます。
'* 注意事項：このテストを行う際は、シート：ブラウザ起動設定 にて、`常にUTF-8でCDP-Json送信`をONにしてください
'***************************************************************************************************
Sub JapaneseElementTest()
    '設定シートに基づくブラウザ立ち上げ、体脂肪率計算サイトへアクセスします
    Dim Demo_Japanese As WebView2Browser: Set Demo_Japanese = 設定シートからのWebView2起動("https://keisan.site/exec/system/1161228728")
    
    
    ' 身長をセット
    Dim height As CDPElement
    Set height = Demo_Japanese.getElementByID("var_身長")
    
    '日本語と絵文字入力テスト
    height.sendString "うみねこ！" & WorksheetFunction.Unichar(128566) & WorksheetFunction.Unichar(8205) & WorksheetFunction.Unichar(127787) & WorksheetFunction.Unichar(65039) & "みゃ～お！" & WorksheetFunction.Unichar(129442)  '日本語兼サロゲートペア絵文字入力テスト(U+1F636 U+200D U+1F32B U+FE0F、U+1F9A2)
    Demo_Japanese.notify "身長を入力しました" & WorksheetFunction.Unichar(129418)       '日本語兼絵文字通知表示テスト(U+1F98A)
    Demo_Japanese.sleep 3

    'ちゃんと数字で入力しなおす
    height.sendString "170.5"
    Demo_Japanese.notify "身長を入力し直しました" & WorksheetFunction.Unichar(128397) & WorksheetFunction.Unichar(65039)    '日本語兼サロゲートペア絵文字通知表示テスト(U+1F58D U+FE0F)
    Demo_Japanese.sleep 3
    
    ' 体重をセット
    Dim weight As CDPElement
    Set weight = Demo_Japanese.getElementByID("var_体重")
    weight.sendString "48.5"
    Demo_Japanese.notify "体重を入力しました" & WorksheetFunction.Unichar(9878) & WorksheetFunction.Unichar(65039)    '日本語兼サロゲートペア絵文字通知表示テスト(U+2696 U+FE0F)
    Demo_Japanese.sleep 3

    ' ボタンクリック
    Demo_Japanese.getElementByID("executebtn").click
    Demo_Japanese.notify "体脂肪率を計算しました" & WorksheetFunction.Unichar(129518)    '日本語兼絵文字通知表示テスト(U+1F9EE)
    Demo_Japanese.sleep 3

    ' 体脂肪率を取得
    Dim 体脂肪率 As Double
    体脂肪率 = Demo_Japanese.getElementByID("ans1").innerText
    Debug.Print "体脂肪率は、" & 体脂肪率 & "% です。"


    'ブラウザを閉じる。demo終了
    Unload WebView2InExcelForm
    Demo_Japanese.quit
End Sub

'***************************************************************************************************
'* 機能　　：JavaScript関数、`alert`処理に関するDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：非同期実行、イベントキャプチャした内容をもとにコマンド実行といったことをデモンストレーションします
'* 注意事項：このライブラリのメソッドは、同期前提で組まれてる都合上、低レベル操作で記述します
'***************************************************************************************************
Sub TestAlert()
    '設定シートに基づくブラウザ立ち上げ。`selenium`の独自テストページに遷移します
    Dim Demo_alerts As WebView2Browser: Set Demo_alerts = 設定シートからのWebView2起動("https://www.selenium.dev/selenium/web/alerts.html")

    'JavaScriptのアラート表示イベントを購読し、閉じれるようにしておきます
    Demo_alerts.SubscribeDevToolsProtocolEvent "Page.javascriptDialogOpening"


    '必要な変数を用意
    Dim paramsCDP As New Scripting.Dictionary
    Dim resCDP As Scripting.Dictionary
    Dim searchId As String
    Dim nodeId As Long
    Dim x As Double, y As Double
    
    'テキスト入力用のAlertに入力させる文字列の指定
    Dim 入力文字内容 As String: 入力文字内容 = "VBAから入力したテスト文字列です！" & WorksheetFunction.Unichar(129418)
    

    With Demo_alerts
        ' --- 1. 必要なドメインを有効化 ---
        .invokeMethod ("DOM.enable")
        .invokeMethod ("Page.enable")
        

        ' --- 2. DOMツリーを同期させ、ID割り振りを行う ---
        paramsCDP.RemoveAll
        paramsCDP.Add "depth", 0        '返却時のDOM情報は不要なので、0にしておく
        paramsCDP.Add "pierce", True    'Shadow DOMの中まで貫通させる
        .invokeMethod "DOM.getDocument", paramsCDP
        ' これでブラウザ内の全ノードにIDが割り振られます

        Dim i As Long
        For i = 1 To 3
            Dim TargetXpath As String
            Select Case i
                Case 1: TargetXpath = "//*[@id='alert']"
                Case 2: TargetXpath = "//*[@id='empty-alert']"
                Case 3: TargetXpath = "//*[@id='prompt']"
            End Select

            ' --- 3. XPathで検索 (Shadow DOMの貫通も可) ---
            paramsCDP.RemoveAll
            paramsCDP.Add "query", TargetXpath  '先頭のリンクを対象に
            Set resCDP = .invokeMethod("DOM.performSearch", paramsCDP)
            searchId = resCDP("searchId")
    
    
            ' --- 4. nodeIdを取得 ---
            paramsCDP.RemoveAll
            paramsCDP.Add "searchId", searchId
            paramsCDP.Add "fromIndex", 0   '先頭の件数から
            paramsCDP.Add "toIndex", 1     '1件分のみ
            Set resCDP = .invokeMethod("DOM.getSearchResults", paramsCDP)
            nodeId = resCDP("nodeIds")(1)  '配列の先頭を取得
    
    
            ' --- 5. nodeId を objectId に変換 ---
            paramsCDP.RemoveAll
            paramsCDP.Add "nodeId", nodeId
            Set resCDP = .invokeMethod("DOM.resolveNode", paramsCDP)


            ' --- 6. 非同期でコマンド準備/実行(Jsのクリック処理) ---
            paramsCDP.RemoveAll
            paramsCDP.Add "objectId", resCDP("object")("objectId")
            paramsCDP.Add "functionDeclaration", "function() { this.click(); }"
            Dim AsyncID As Long

            ' --- 7. イベントキャプチャを有効化 ---
            Set .BrowserEvents = New Dictionary

            'この瞬間、JavaScriptの`alert`関数が発動され、ダイアログOpenイベントも発動されます
            AsyncID = .invokeMethodAsync("Runtime.callFunctionOn", paramsCDP, alwaysBrowserContext:=False)
    
            ' --- 8. 特定のイベント名が出るまでループ ---
            Const SearchEventName As String = "Page.javascriptDialogOpening"
            Do
                '非同期イベントを取り出す
                .TakeEvents
                DoEvents
    
                'イベント名の確認
                If .BrowserEvents("EventMethods").Exists(SearchEventName) Then
                    '出ているダイアログの情報の確認
                    Dim tmp
                    For Each tmp In .BrowserEvents("EventMethods")(SearchEventName)
                        Debug.Print "url    :"; tmp("params")("url")
                        Debug.Print "message:"; tmp("params")("message")
                        Debug.Print "type   :"; tmp("params")("type") & vbCrLf
                    Next
    
                    '見つかったので抜ける
                    Exit Do
                End If
            Loop While True
    
    
            ' --- 9. ダイアログに反応しておく ---
            paramsCDP.RemoveAll
            paramsCDP.Add "accept", True
            paramsCDP.Add "promptText", 入力文字内容
            Set resCDP = .invokeMethod("Page.handleJavaScriptDialog", paramsCDP)
    
    
            ' --- 10. 以前、非同期で実行した結果も拝見する ---
'            Dim ErrorExist As Boolean
'            Dim resCDPAsync As Dictionary
'            Dim jsonconv As New WebJsonConverter
'            Set resCDPAsync = .ResultCDPForAsync(AsyncID, ErrorExist)
'            If Not (ErrorExist) Then Debug.Print "resCDPAsync - " & jsonconv.ConvertToJson(resCDPAsync)
        Next


        ' --- 11. ブラウザを閉じる ---
        Dim Htmlの表示内容 As String: Htmlの表示内容 = .getElementByXPath("//*[@id='text']/p").innerText
        Debug.Print "htmlの出力文字列：" & Htmlの表示内容
        Debug.Assert Htmlの表示内容 = 入力文字内容
        
        Unload WebView2InExcelForm
        .quit
    End With
End Sub

'***************************************************************************************************
'* 機能　　：拡張機能を読み込むDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ブラウザ自身をターゲットとした`invokeMethod`の使用例です
'* 注意事項：・`Extensions`は実験的ドメインですが、Class内Err.Raiseでは止めずに、ここの自力判定でエラーハンドリングします
'            ・Demoコードに載せておきながら、機能しません。Excel上で読み込むWebView2の場合は制限があるようです...
'***************************************************************************************************
Sub UseExtensions()
    '必要な変換オブジェクトを用意
    Dim JsonDicObj As New WebJsonConverter
    
    '拡張機能があるアンパックフォルダパスを、ダイアログで指定
    '参考 → https://qiita.com/studio_haneya/items/9f5141b667efc3bfa615
    Dim ExtensionsFolderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "拡張機能の基となる`manifest.json`を含むフォルダを選択してください"
        .InitialFileName = Environ("UserProfile") & "\AppData\Local"    '初期位置

        If .show = -1 Then ExtensionsFolderPath = .SelectedItems(1) Else Exit Sub
    End With



    '設定シートに基づくブラウザ立ち上げ
    Dim controlExtensions As WebView2Browser: Set controlExtensions = 設定シートからのWebView2起動
    
    '拡張機能のページへ遷移
    controlExtensions.navigate "edge://extensions/"

    '拡張機能を読み込む
    Dim CDPParams As Dictionary, ResultCDP As Dictionary
    Set CDPParams = New Dictionary
    CDPParams.Add "path", ExtensionsFolderPath
    Set ResultCDP = controlExtensions.invokeMethod("Extensions.loadUnpacked", CDPParams, True, False)   '今回は、エラー無視で設定

    '読み込まれたか確認する
    '※コマンド実行に失敗すると、`nothing`で返るので、この仕様を利用します
    If ResultCDP Is Nothing Then
        'CDP-Json結果に`error`要素あり
        MsgBox "拡張機能のインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.LastCDPJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.LastCDPJsonError("code")

        'ブラウザを閉じる。demo終了
        controlExtensions.quit
        Exit Sub

    ElseIf ResultCDP.Exists("id") Then
        MsgBox "拡張機能のインストールに成功しました。ブラウザをご確認ください。" & vbCrLf & "なお、OKを押すと、アンインストールします。", vbInformation, "ExtensionsID：" & ResultCDP("id")
    
    Else
        MsgBox "インストールIDの確認が取れませんでした。" & vbCrLf & vbCrLf & "<RawResult>" & vbCrLf & JsonDicObj.ConvertToJson(ResultCDP), vbExclamation, "Not found id"

        'ブラウザを閉じる。demo終了
        controlExtensions.quit
    End If


    '拡張機能をアンインストール
    Set CDPParams = New Dictionary
    CDPParams.Add "id", ResultCDP("id")
    Set ResultCDP = controlExtensions.invokeMethod("Extensions.uninstall", CDPParams, True, False)

    '消えたか確認する
    If ResultCDP Is Nothing Then
        'CDP-Json結果に`error`要素あり
        MsgBox "拡張機能のアンインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.LastCDPJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.LastCDPJsonError("code")

    Else
        MsgBox "拡張機能のアンインストールに成功しました。ブラウザをご確認ください。", vbInformation, "Uninstall Done!"
    End If


    'ブラウザを閉じる。demo終了
    controlExtensions.quit
End Sub

Sub runEdge()
'------------------------------------------------------
' This is an example of how to use the browser classes
' This demo tries to access a webpage of a famous movie
' and retrieve its current view count.
'------------------------------------------------------
 
   'Start Browser
   'If no browser name is indicated, chrome is started by default.
   'Homepage has been disabled to speed up by default.
   'To skip cleaning active sessions, set cleanActive to False.
   'This will make browser starts faster but at the risk of pipe error if
   'there are other chrome instances already running.
   'If reAttach = False, .start will not automatically try to reattach
   'to previous instances open by CDP but will start a brand new instead.
    Dim edge As WebView2Browser
    Set edge = 設定シートからのWebView2起動
 
   'Navigate and wait
   'If till argument is omitted, will by default wait until ReadyState = complete
    edge.navigate "https://livingwaters.com/movie/the-atheist-delusion/", isInteractive
 
   'Get view count via the new notify method
    Dim viewCount As Long
    viewCount = edge.getElementByQuery("[data-id='4b9a4b19']").innerText
    edge.notify "This free movie has already reached " & viewCount & " views! Wow!"
 
End Sub

Sub runHiddenForJapan()
'---------------------------------------------------------------------------------
' Demonstrate background running of an automated session.
' This demo will try to open Google in the background, then search for an article
' of CodeProject and retrieve its vote count. Once done, it will prompt a message
' to display the browser window.
' It is recommended to make Immediate Window visible so that you can see the
' activity that is running in the background.
' To confirm the result, you can perform the following steps:
'   1. Go to Google.com
'   2. Type "automate edge vba" and click Search
'   3. Click on the first result to reach
'
' ※日本国向けに改良します。
'---------------------------------------------------------------------------------

    Dim chrome As WebView2Browser

    'Start and hide
    Set chrome = 設定シートからのWebView2起動
    chrome.hide

    'Perform automation in the background
    chrome.navigate "https://google.com", isInteractive
    chrome.getElementByQuery("[name='q']").value = "automate edge vba"
    chrome.getElementByQuery("[name='q']").submit

    'Click the target result link
    chrome.getElementByXPath("//h3[text()='Chrome DevTools ProtocolでEdgeを操作するVBAマクロ']").click      '2026/02/16 時点での、最上位結果

    'Confirm result and display
    Dim userChoice As Long
    userChoice = MsgBox("Automation completed. Do you want to see the window?", vbYesNo)
    If userChoice = vbYes Then chrome.show Else chrome.quit

End Sub

Sub runTabsAsOne()
'--------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' Similar to the runInstances example but this is with multiple tabs in
' the same instance instead.
'--------------------------------------------------------------------------
 
    Dim chrome As WebView2Browser
    Set chrome = 設定シートからのWebView2起動
    chrome.show
    
   'Automate Tabs
    chrome.Url = "https://google.com"   'or [chrome.navigate "https://google.com"]
    chrome.newTab "https://sg.yahoo.com"
    chrome.newTab "https://bing.com"
 
   'Resize to complete
    chrome.sleep    'ちょこっとクールタイムが必要みたい
    chrome.show xywh:="0 20 1000 700"
 
End Sub

Sub runTabsAsMany()
'-------------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' This is like having 3 automation instances running together like runInstances.
' However, each tab will have to share the same start settings, unlike
' the case of runInstances where each instance can be setup with a different
' settings to each other.
'-------------------------------------------------------------------------------
 
    Dim chrome As WebView2Browser
    Set chrome = 設定シートからのWebView2起動
    chrome.show
 
   'Create and assign tabs
    Dim tab1 As New WebView2Browser                   'The keyword "New" is a must
    Dim tab2 As New WebView2Browser
    Dim tab3 As New WebView2Browser
    Set tab1 = chrome                            'The first tab is open by default after .start
    Set tab2 = chrome.newTab    'newWindow: open tab as a new window instead of a tab
    Set tab3 = chrome.newTab
 
   'Automate each tabs
    tab1.navigate "https://google.com"
    tab2.navigate "https://sg.yahoo.com"
    tab3.navigate "https://bing.com"
 
   'Resize to complete
    tab1.show xywh:="0 10 1000 700"
    tab2.show xywh:="0 45 1000 700"
    tab3.show xywh:="0 90 1000 700"
 
End Sub

Sub runNewTab()
'--------------------------------------------------------------------------
' This example demonstrates:
' 1. The use of advanced arguments feature added by Long Vh to
'    allow the choice of additional settings for the automation pipe. See
'    https://peter.sh/experiments/chromium-command-line-switches/
' 2. The xPath technique to directly modify the current HTML element
'    so that it will behave in a new way that it was not so before.
' 3. The technique employed to integrate the new tab open spontaneously
'    by interaction with the webpage (instead of using .newTab) into the
'    automation pipe for further processing on the new tab.
'--------------------------------------------------------------------------
 
   'Init browser with custom arguments
    Dim chrome As WebView2Browser
    Set chrome = 設定シートからのWebView2起動
    'chrome.start addArgs:="--disable-popup-blocking"    'The disable-popup-blocking argument is needed to allow opening link in a new tab
    'chrome.show asMaximized
    
   'Perform standard google search
    chrome.navigate "https://google.com"
    chrome.getElementByQuery("[name='q']").value = "newstarget.com"
    chrome.getElementByQuery("[name='q']").submit

   'Google search result returns links that open in the same tab window
   'For this demonstration, we need to make it open in a new tab window instead
    Dim targetElement As CDPElement
    Set targetElement = chrome.getElementByXPath(".//a[contains(@href, 'https://www.newstarget.com/')]")
    targetElement.setAttribute "target", "_blank"   'Modify the element attribute to open in a new tab instead
    targetElement.click                             'Click the link, a new tab will be spontaneously open
    
   'Use getTabNew to quickly refer to the next newly open tab
    Dim targetTab As New WebView2Browser
    Set targetTab = chrome.getTab(Url:="https://www.newstarget.com/")
    targetTab.wait
 
   'Feed the top news title for today
    Dim firstTitle As String
    firstTitle = targetTab.getElementByQuery("div[class='Headline']").innerText
    targetTab.notify "Top popular headline for the day is """ & firstTitle & """."
    targetTab.printTabs True
    Debug.Print targetTab.CurrentTargetID
 
End Sub



'----------------------------------------------------------------------
' TestWebView2Simple
'   最もシンプルなデモ。Excelウィンドウに WebView2 を重ねて表示。
'   ※実際のアプリでは UserForm の Frame の hWnd を使う
'----------------------------------------------------------------------
Public Sub TestWebView2Simple()
    Dim wv2 As New WebView2Core
Dim hwndParent
    ' Excel のメインウィンドウハンドルを取得
    Dim hWnd As LongPtr
    hWnd = Application.hWnd

    ' 画面左上 300x400px に WebView2 を表示
    Dim ok As Boolean
    ok = wv2.Initialize(hWnd, 0, 0, 800, 600, "https://eschamali.github.io/StarterWebScrapingKit/")

    If Not ok Then
        MsgBox "初期化コマンド送信失敗: " & wv2.LastErrorDescription, vbCritical
        Exit Sub
    End If

    ' ---- 初期化完了待機ループ ----
    ' DoEvents だけでは COM STA コールバックが層かない場合があるため、
    ' ProcessMessages（PeekMessage/DispatchMessage）でメッセージキューを層かせる
    Debug.Print "[WV2] Waiting for Ready... (check Immediate Window for callback logs)"
    Dim t As Single: t = Timer
    Do While Not wv2.IsReady And Timer - t < 15
        wv2.ProcessMessages
    Loop

    If wv2.IsReady Then
        MsgBox "WebView2 の初期化に成功しました！" & vbCrLf & "OKを押すと閉じます。", vbInformation
    Else
        MsgBox "タイムアウト。LastError: 0x" & Hex(wv2.LastErrorCode) & vbCrLf & wv2.LastErrorDescription, vbCritical
    End If

    wv2.quit
    Set wv2 = Nothing
End Sub

'----------------------------------------------------------------------
' TestWebView2Form  ★推奨★
'   UserForm の Frame hWnd を親として WebView2 を埋め込むデモ。
'   vbModeless なので Excel の操作を維持しながら使える。
'
'   ★ Application.hWnd を親にするとクラッシュする問題の解決版 ★
'----------------------------------------------------------------------
Public Sub TestWebView2Form()
    WebView2InExcelForm.show vbModeless
End Sub

'----------------------------------------------------------------------
' TestWebView2FormModal
'   モーダル版（Excel 操作をブロックして WebView2 を表示）
'----------------------------------------------------------------------
Public Sub TestWebView2FormModal()
    WebView2InExcelForm.show
End Sub

'----------------------------------------------------------------------
' TestWebView2CDP ? CDP (CallDevToolsProtocolMethod) の動作確認
' Immediate に [CDP] Invoke が出ればコールバックは呼ばれている。出なければ OFF_WV2_CallDevToolsProtocolMethod の index を 22,36,37,39 などに変更して再試行
'----------------------------------------------------------------------
Public Sub TestWebView2CDP()
    Dim wv2 As New WebView2Core
    Dim hWnd As LongPtr: hWnd = Application.hWnd
    If Not wv2.Initialize(hWnd, 0, 0, 600, 400, "about:blank", "Automation Data") Then
        MsgBox "Initialize 失敗: " & wv2.LastErrorDescription, vbCritical
        Exit Sub
    End If
    Dim t As Single: t = Timer
    Do While Not wv2.IsReady And Timer - t < 15
        wv2.ProcessMessages
    Loop
    If Not wv2.IsReady Then
        MsgBox "Ready タイムアウト", vbCritical
        wv2.quit
        Exit Sub
    End If
    Dim params As String: params = "{""expression"":""1+1""}"
    Dim result As String: result = wv2.CallDevToolsProtocolMethod("Runtime.evaluate", params)
    Debug.Print "[CDP] Runtime.evaluate result: " & result
    MsgBox "CDP 結果（Immediate を確認）: " & Left$(result, 200) & IIf(Len(result) > 200, "...", ""), vbInformation
    wv2.quit
    Set wv2 = Nothing
End Sub


Sub WebView2Browserクラスで起動するDemo()
    Dim test As WebView2Browser: Set test = 設定シートからのWebView2起動("http://officetanaka.net/")
    Dim paramCDP As New Dictionary, ResultCDP As New Dictionary
    paramCDP.Add "expression", "1+1"
    Set ResultCDP = test.invokeMethod("Runtime.evaluate", paramCDP)
    Dim JsonConv As New WebJsonConverter

    MsgBox "CDP 結果: " & JsonConv.ConvertToJson(ResultCDP), vbInformation
    
    test.navigate "https://eschamali.github.io/StarterWebScrapingKit/#userform-summary"

    test.quit
    
End Sub

Sub reAttachDemo()
    Dim test As WebView2Browser: Set test = 設定シートからのWebView2起動("http://officetanaka.net/")
    Dim paramCDP As New Dictionary, ResultCDP As New Dictionary
    paramCDP.Add "expression", "1+1"
    Set ResultCDP = test.invokeMethod("Runtime.evaluate", paramCDP)
    Dim JsonConv As New WebJsonConverter

    MsgBox "CDP 結果: " & JsonConv.ConvertToJson(ResultCDP), vbInformation
End Sub


Sub 再開プロシージャ()
    Dim test As New WebView2Browser
    test.reattach "Automation Data"  '何らかの方法で`WebView2Browserクラスで起動するDemo`で実行した情報を復元注入
    test.navigate "https://eschamali.github.io/StarterWebScrapingKit/#userform-summary"

    test.quit





End Sub





Sub Demo_NetworkResponseReceived()
    Dim core As WebView2Browser: Set core = 設定シートからのWebView2起動

    ' --- イベントキャプチャ開始 ---
    Set core.BrowserEvents = New Dictionary

    ' DevToolsProtocol event receiver を購読（WebView2側イベントはここで受けます）
    If Not core.SubscribeDevToolsProtocolEvent("Network.responseReceived") Then
        Debug.Print "[WV2][CDP] SubscribeDevToolsProtocolEvent failed. (可能性: vtable index 調整が必要)"
    End If

    ' Network.enable（同期でOK）
    core.invokeMethod "Network.enable"

    ' 遷移（ここからイベントが飛んでくる）
    core.navigate "http://officetanaka.net/"

    ' 取り出しの重複防止用（__index__ で追跡）
    Dim lastSeenIndex As Long: lastSeenIndex = 0
    Dim CharConv As New CharacterCodeConversion
    Dim JsonConv As New WebJsonConverter

    ' requestId を収集し、後段でまとめて Network.getResponseBody を叩く
    ' （イベント処理ループ内で同期CDPを重ねると、ブレークポイント等でタイミングが崩れて不安定になりやすいため）
    Dim reqQueue As Object: Set reqQueue = CreateObject("Scripting.Dictionary") ' key=requestId, value= "url|status"

    Dim stopAt As Single: stopAt = Timer + 30
    Do While Timer < stopAt
        core.ProcessMessages 200   ' COMコールバック配送
        core.TakeEvents           ' キューを BrowserEvents に反映

        If Not (core.BrowserEvents Is Nothing) Then
            Dim evMethods As Dictionary
            Set evMethods = core.BrowserEvents("EventMethods")

            If evMethods.Exists("Network.responseReceived") Then
                CharConv.BytesToSaveFile JsonConv.ConvertToJson(core.BrowserEvents), "C:\Users\XXX\Downloads\test", "WebView2からの非同期イベント情報.json"
                
                Dim lst As Collection
                Set lst = evMethods("Network.responseReceived")

                Dim i As Long
                For i = 1 To lst.Count
                    Dim ev As Dictionary
                    Set ev = lst(i)

                    Dim idx As Long: idx = ev("__index__")
                    If idx > lastSeenIndex Then
                        lastSeenIndex = idx

                        ' responseReceived は params.requestId が欲しい
                        Dim p As Dictionary: Set p = ev("params")
                        Dim requestId As String: requestId = p("requestId")

                        Dim response As Dictionary
                        Set response = p("response")

                        Debug.Print "response: url=" & response("url") & ", status=" & response("status")

                        ' ここでは requestId を収集するだけ（body取得は後段）
                        If Not reqQueue.Exists(requestId) Then
                            reqQueue.Add requestId, CStr(response("url")) & "|" & CStr(response("status"))
                        End If
                    End If
                Next i
            End If
        End If

        DoEvents
    Loop

    ' ---- 検証: body取得フェーズに入る前にイベント購読を解除 ----
    '   以降の処理では Network.responseReceived が追加で積まれない状態になる想定
    core.UnsubscribeDevToolsProtocolEvent

    ' ---- 収集した requestId から body をまとめて取得（ブレークポイントを当てても比較的安定） ----
    Dim k As Variant
    For Each k In reqQueue.keys
        Dim paramsJson As New Dictionary
        paramsJson.Add "requestId", CStr(k)

        Dim bodyRes As Dictionary
        ' 完走優先：StopError=False（内部でNothingが返る可能性あり）
        Set bodyRes = core.invokeMethod("Network.getResponseBody", paramsJson, , False)
        If Not bodyRes Is Nothing Then
            If bodyRes.Exists("body") Then
                Dim Body As String: Body = bodyRes("body")
                Debug.Print "getResponseBody: requestId=" & CStr(k) & " body sample=" & Left$(Body, 200)
            Else
                Debug.Print "getResponseBody: requestId=" & CStr(k) & " (no body field)"
            End If
        Else
            Debug.Print "getResponseBody: requestId=" & CStr(k) & " -> Nothing"
        End If
        DoEvents
    Next k

    Unload WebView2InExcelForm
    core.quit
End Sub





