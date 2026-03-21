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
