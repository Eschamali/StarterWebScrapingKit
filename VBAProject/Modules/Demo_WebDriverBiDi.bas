Attribute VB_Name = "Demo_WebDriverBiDi"
'==============================================================================================================
'               Automating Chromium-Based Browsers with WebDriverBiDi API and VBA
'--------------------------------------------------------------------------------------------------------------
'
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                               ■■■ 設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：設定シートから、パラメーターを読み込んで、BiDiモードでブラウザを起動するヘルパープロシージャです
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - WebDriverBiDiCore
'* 引数　　：StartURL                       ブラウザ起動時にアクセスしたいURL。指定しない場合は、空ページ(abount:blank)になります。
'            SwtchUser                      マルチインスタンス用に別ユーザーを指定するときに使用します
'            KioskMode                      0(省略)：通常モード(キオスクモードは使いません)
'                                           1      ：キオスクモード デジタル/対話型サイネージ
'                                           2      ：キオスクモード パブリック ブラウジング
'
'            sessionCapabilitiesRequest     `session.new`のParametersをセットします。予めDictionaryで組み立ててください
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBEによるハードコーディングではなく、設定シートから読み込む方式により、ユーザー側からも手軽に設定変更ができます
'* 注意事項：Demoモジュールにあるコードですが、他の部分で共用してるため、消さずにどこかにカット&ペーストしておくとよいでしょう
'***************************************************************************************************
Public Function 設定シートからのBiDi起動(Optional StartURL As String, Optional SwitchUser As String, Optional KioskMode As edgeKioskType, Optional sessionCapabilitiesRequest As Dictionary) As WebDriverBiDiCore
    '設定シートの各セルから設定値を取得し、適用
    With ShSetting01_StartBrowser
        '起動ブラウザ種類の設定
        '※BiDi-Json コマンドによる操作ですが、Chromium系統に特化した制御のため、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim ブラウザ名 As String: ブラウザ名 = IIf(.Range(.UseRangeName(4, "Demo_WebDriverBiDi.設定シートからのBiDi起動")).value, "chrome", "edge")

        '第2引数が省略ならシート側の設定を適用
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .Range(.UseRangeName(2, "Demo_WebDriverBiDi.設定シートからのBiDi起動")).value, SwitchUser)

        'ブラウザ起動
        Set 設定シートからのBiDi起動 = New WebDriverBiDiCore
        設定シートからのBiDi起動.start ブラウザ名, StartURL, .Range(.UseRangeName(6, "Demo_WebDriverBiDi.設定シートからのBiDi起動")).value, UseDataDir, .Range(.UseRangeName(3, "Demo_WebDriverBiDi.設定シートからのBiDi起動")).value, KioskMode, sessionCapabilitiesRequest
    End With
End Function

Sub BiDiによる冒険の始まり()
    '設定シートに基づくブラウザ立ち上げ
    Dim HelloWorldAutomationBrowser As WebDriverBiDiCore: Set HelloWorldAutomationBrowser = 設定シートからのBiDi起動

    '↓ここから、あなたのイメージをコードに落とし込む↓




    'ブラウザを正常に閉じる
    HelloWorldAutomationBrowser.quit
End Sub



'***************************************************************************************************
'                               ■■■ Demoプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：イベントキャプチャに関するDemoコード(BiDi版)です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：CDP版の「ネットワークイベントの確認」をBiDiの`network`ドメインを用いて再現したデモです。
'*           `session.subscribe` で `network` 関連イベントを購読し、結果をJSON出力します。
'***************************************************************************************************
Sub ネットワークイベントの確認()
    '必要な変換オブジェクトを用意
    Dim JsonDicObj As New WebJsonConverter
    Dim CharConvObj As New CharacterCodeConversion
    
    'WebDriverBiDiCoreの初期化とブラウザ立ち上げ
    Dim Demo_NetworkEvent As WebDriverBiDiCore
    Set Demo_NetworkEvent = 設定シートからのBiDi起動
    
    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary
    
    '現在のコンテキストIDを取得する (ここではざっくり1番目のコンテキストを利用)
    Set resultBiDi = Demo_NetworkEvent.invokeMethod("browsingContext.getTree")
    Dim targetContext As String
    If Not (resultBiDi Is Nothing) Then
        targetContext = resultBiDi("contexts")(1)("context")
    End If
    
    '-------------------------------- 機能1：イベントキャプチャを有効化する --------------------------------
    '`New Dictionary`を渡すことで、内部で非同期イベントの蓄積を開始する
    Set Demo_NetworkEvent.BiDiEvents = New Dictionary

    'BiDi側でネットワークイベントを購読開始する
    Set paramsBiDi = New Dictionary
    Dim eventsArray As New Collection
    eventsArray.Add "network.beforeRequestSent"
    eventsArray.Add "network.responseCompleted"
    paramsBiDi.Add "events", eventsArray
    
    Demo_NetworkEvent.invokeMethod "session.subscribe", paramsBiDi
    
    'URL遷移して、読み込み終わるまで待機
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "context", targetContext
    paramsBiDi.Add "url", "http://officetanaka.net/excel/vba/file/file11.htm"
    paramsBiDi.Add "wait", "complete"
    Demo_NetworkEvent.invokeMethod "browsingContext.navigate", paramsBiDi

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う (念のため待機後にも余波を回収)
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_Event.json"


    '-------------------------------- 機能2：セーブデータを作成し、イベントキャプチャを無効化する --------------------------------
    Dim SaveDataEvents As Dictionary: Set SaveDataEvents = Demo_NetworkEvent.BiDiEvents  'セーブデータ作成
    Set Demo_NetworkEvent.BiDiEvents = Nothing               '`Nothing`を渡すことで、イベント記録状態を破棄する


    'URL遷移
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "context", targetContext
    paramsBiDi.Add "url", "http://officetanaka.net/youtube/20200714b.htm"
    paramsBiDi.Add "wait", "complete"
    Demo_NetworkEvent.invokeMethod "browsingContext.navigate", paramsBiDi

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存しますが、無効中なので破棄状態（0バイト等）になります
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_NotEvent.json"


    '-------------------------------- 機能3：セーブデータを読み込み、そこからイベントキャプチャを再開する --------------------------------
    Set Demo_NetworkEvent.BiDiEvents = SaveDataEvents        '既存のセーブデータを読み込む
    
    'URL遷移
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "context", targetContext
    paramsBiDi.Add "url", "http://officetanaka.net/index.stm"
    paramsBiDi.Add "wait", "complete"
    Demo_NetworkEvent.invokeMethod "browsingContext.navigate", paramsBiDi

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_EventFromSaveData.json"


    'ブラウザを閉じる。demo終了
    Demo_NetworkEvent.quit
End Sub

'***************************************************************************************************
'* 機能　　：拡張機能を読み込むDemoコード(BiDi版)です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：BiDiプロトコルの `webExtension` モジュールを使用した拡張機能のインストール・アンインストールのデモです。
'* 注意事項：・このテストを行う際は、事前シート：ブラウザ起動設定 にて、`CDP-Jsonで拡張機能を制御` をONにしてください
'***************************************************************************************************
Sub UseExtensions()
    '必要な変換オブジェクトを用意
    Dim JsonDicObj As New WebJsonConverter
    
    '拡張機能があるアンパックフォルダパスを、ダイアログで指定
    Dim ExtensionsFolderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "拡張機能の基となる`manifest.json`を含むフォルダを選択してください"
        .InitialFileName = Environ("UserProfile") & "\AppData\Local"    '初期位置

        If .show = -1 Then ExtensionsFolderPath = .SelectedItems(1) Else Exit Sub
    End With

    'WebDriverBiDiCoreの初期化とブラウザ立ち上げ
    Dim controlExtensions As WebDriverBiDiCore
    
    '---- 拡張機能制御を有効化するオプションを作成 ---
    Dim caps As New Dictionary
    Dim alwaysMatch As New Dictionary
    
    ' BiDiでは、セッション確立時の引数として渡すか、WebDriver側のCapabilityで有効にする必要がありますが、
    ' CDPBrowserの仕組み（引数渡し）を利用するためそのまま起動します。
    caps.Add "capabilities", New Dictionary
    caps("capabilities").Add "alwaysMatch", alwaysMatch
    '-------------------------------------------------

    ' 起動
    Set controlExtensions = 設定シートからのBiDi起動(sessionCapabilitiesRequest:=caps)

    '現在のコンテキストIDを取得する
    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary
    Set resultBiDi = controlExtensions.invokeMethod("browsingContext.getTree")
    Dim targetContext As String
    If Not (resultBiDi Is Nothing) Then
        targetContext = resultBiDi("contexts")(1)("context")    '一旦は、先頭タブで　※本来はURLcheckとかがいると思うが、低レベル制御の都合上、妥協
    End If

    '拡張機能のテストページ（もしくは任意のページ）へ遷移
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "context", targetContext
    paramsBiDi.Add "url", "edge://extensions/"
    paramsBiDi.Add "wait", "complete"
    controlExtensions.invokeMethod "browsingContext.navigate", paramsBiDi

    '-----------------------------------------------------------------------
    '拡張機能を読み込む (BiDi `webExtension.install`)
    '-----------------------------------------------------------------------
    Dim extData As New Dictionary
    extData.Add "type", "path"
    extData.Add "path", ExtensionsFolderPath
    paramsBiDi.Add "extensionData", extData
    
    ' 今回はエラー無視で設定 (StopError:=False)
    Set resultBiDi = controlExtensions.invokeMethod("webExtension.install", paramsBiDi, False)

    '読み込まれたか確認する
    If resultBiDi Is Nothing Then
        ' コマンド実行に失敗した場合、LastBiDiJsonError からエラー情報を取得する
        MsgBox "拡張機能のインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.LastBiDiJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.LastBiDiJsonError("error")

        'ブラウザを閉じる。demo終了
        controlExtensions.quit
        Exit Sub

    ElseIf resultBiDi.Exists("extension") Then
        ' BiDiの webExtension.install は `extension` キーで IDを返します。
        MsgBox "拡張機能のインストールに成功しました。ブラウザをご確認ください。" & vbCrLf & "なお、OKを押すと、アンインストールします。", vbInformation, "ExtensionsID：" & resultBiDi("extension")
    
    Else
        MsgBox "インストールIDの確認が取れませんでした。" & vbCrLf & vbCrLf & "<RawResult>" & vbCrLf & JsonDicObj.ConvertToJson(resultBiDi), vbExclamation, "Not found id"

        'ブラウザを閉じる。demo終了
        controlExtensions.quit
        Exit Sub
    End If

    '-----------------------------------------------------------------------
    '拡張機能をアンインストール (BiDi `webExtension.uninstall`)
    '-----------------------------------------------------------------------
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "extension", resultBiDi("extension")
    Set resultBiDi = controlExtensions.invokeMethod("webExtension.uninstall", paramsBiDi, False)

    '消えたか確認する
    If resultBiDi Is Nothing Then
        MsgBox "拡張機能のアンインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.LastBiDiJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.LastBiDiJsonError("error")
    Else
        MsgBox "拡張機能のアンインストールに成功しました。ブラウザをご確認ください。", vbInformation, "Uninstall Done!"
    End If

    'ブラウザを閉じる。demo終了
    controlExtensions.quit
End Sub

'***************************************************************************************************
'* 機能　　：JavaScript関数、`alert`処理に関するBiDi版のDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：非同期実行、イベントキャプチャした内容をもとにコマンド実行といったことをデモンストレーションします
'***************************************************************************************************
Sub TestAlert()
    '必要な変換オブジェクトを用意
    Dim JsonDicObj As New WebJsonConverter

    'WebDriverBiDiCoreの初期化とブラウザ立ち上げ
    Dim Demo_alerts As New WebDriverBiDiCore
    
    '---- JavaScriptによる自動アラート処理を無効化するオプションを作成 ---
    Dim caps As New Dictionary
    
    Dim alwaysMatch As New Dictionary
    alwaysMatch.Add "unhandledPromptBehavior", "ignore"
    
    caps.Add "capabilities", New Dictionary
    caps("capabilities").Add "alwaysMatch", alwaysMatch
    '---------------------------------------------------------------------

    'オプションを適用させて、指定URLから直接起動
    Set Demo_alerts = 設定シートからのBiDi起動("https://www.selenium.dev/selenium/web/alerts.html", sessionCapabilitiesRequest:=caps)

    '結果とBiDiパラメーター変数を用意
    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary

    '現在のコンテキストIDを取得する
    Set resultBiDi = Demo_alerts.invokeMethod("browsingContext.getTree")
    Dim targetContext As String
    If Not (resultBiDi Is Nothing) Then
        targetContext = resultBiDi("contexts")(1)("context")    '一旦は、先頭ブラウザタブで　※本来はURLcheckとかがいると思うが、低レベル制御の都合上、妥協
    End If

    'テスト入力文字列
    Dim 入力文字内容 As String: 入力文字内容 = "VBAから入力したテスト文字列です！" & WorksheetFunction.Unichar(129418)
    
    With Demo_alerts
        ' --- 1. 必要なドメイン(イベント)をサブスクライブ ---
        Set paramsBiDi = New Dictionary
        Dim eventsArray As New Collection
        eventsArray.Add "browsingContext.userPromptOpened"
        paramsBiDi.Add "events", eventsArray
        .invokeMethod "session.subscribe", paramsBiDi
        
        Dim i As Long
        For i = 1 To 3
            Dim targetID As String
            Select Case i
                Case 1: targetID = "alert"
                Case 2: targetID = "empty-alert"
                Case 3: targetID = "prompt"
            End Select

            ' --- 2. イベントキャプチャを新しく有効化 ---
            ' 過去のイベントをリセット
            Set .BiDiEvents = New Dictionary
            
            ' --- 3. 非同期でコマンド準備/実行(Jsのクリック処理) ---
            ' 対象の要素をクリックするJSを評価する
            Set paramsBiDi = New Dictionary
            paramsBiDi.Add "expression", "document.getElementById('" & targetID & "').click()"
            Dim targetDict As Dictionary
            Set targetDict = New Dictionary
            targetDict.Add "context", targetContext
            paramsBiDi.Add "target", targetDict
            paramsBiDi.Add "awaitPromise", False
            
            Dim AsyncID As Long
            'この瞬間、JavaScriptの`alert`関数が非同期で発動されます
            AsyncID = .invokeMethodAsync("script.evaluate", paramsBiDi)
    
            ' --- 4. 特定のイベント名が出るまでループ ---
            Const SearchEventName As String = "browsingContext.userPromptOpened"
            Do
                '非同期イベントを取り出す
                .TakeEvents
    
                'イベント名の確認
                If .BiDiEvents("EventMethods").Exists(SearchEventName) Then
                    '出ているダイアログの情報の確認
                    Dim tmp
                    For Each tmp In .BiDiEvents("EventMethods")(SearchEventName)
                        Debug.Print "message:"; tmp("params")("message")
                        Debug.Print "type   :"; tmp("type") & vbCrLf
                    Next
    
                    '見つかったので抜ける
                    Exit Do
                End If
            Loop While True
    
            ' --- 5. ダイアログに反応しておく ---
            Set paramsBiDi = New Dictionary
            paramsBiDi.Add "context", targetContext
            paramsBiDi.Add "accept", True
            paramsBiDi.Add "userText", 入力文字内容
            Set resultBiDi = .invokeMethod("browsingContext.handleUserPrompt", paramsBiDi)
    
            ' --- 6. 以前、非同期で実行した結果も拝見する ---
            Dim resBiDiAsync As Dictionary
            .sleep 0.5 ' 結果取得のためのディレイ
            .TakeEvents ' 受信キューを消化
            
            Dim エラー確認 As Boolean
            Set resBiDiAsync = .ResultBiDiForAsync(AsyncID, エラー確認)
            If Not (resBiDiAsync Is Nothing) Then Debug.Print "resBiDiAsync - " & JsonDicObj.ConvertToJson(resBiDiAsync)
            
        Next

        ' --- 7. ブラウザを閉じる ---
        ' DOM経由のテキスト取得を、script.evaluateで代替
        Set paramsBiDi = New Dictionary
        paramsBiDi.Add "expression", "document.querySelector('#text > p') ? document.querySelector('#text > p').innerText : 'Not Found'"
        Set targetDict = New Dictionary
        targetDict.Add "context", targetContext
        paramsBiDi.Add "target", targetDict
        paramsBiDi.Add "awaitPromise", True
        Set resultBiDi = .invokeMethod("script.evaluate", paramsBiDi)
        
        Dim Htmlの表示内容 As String
        If Not (resultBiDi Is Nothing) Then
            If resultBiDi.Exists("result") Then
                If resultBiDi("result").Exists("value") Then Htmlの表示内容 = resultBiDi("result")("value")
            End If
        End If
        
        Debug.Print "htmlの出力文字列：" & Htmlの表示内容
        Debug.Assert Htmlの表示内容 = 入力文字内容
        .quit
    End With
End Sub

'***************************************************************************************************
'* 機能　　：BiDi+ (Chromium独自拡張) の `goog:cdp.sendCommand` を試すDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebDriver BiDi プロトコルにまだ存在しない詳細な機能を、従来のCDPコマンドを
'*           トンネリング（中継）して呼び出す「BiDi+」の機能デモンストレーションです。
'***************************************************************************************************
Sub TestBiDiPlus_CDPTunnel()
    Dim JsonDicObj As New WebJsonConverter
    Dim bidiPlus As WebDriverBiDiCore
    
    ' ブラウザ起動
    Set bidiPlus = 設定シートからのBiDi起動

    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary

    '-----------------------------------------------------------------------
    ' 1. CDPのセッションIDを取得する (goog:cdp.getSession)
    '-----------------------------------------------------------------------
    ' まず現在のBiDiコンテキストを取得
    Dim targetContext As String
    Set resultBiDi = bidiPlus.invokeMethod("browsingContext.getTree")
    If Not resultBiDi Is Nothing Then
        targetContext = resultBiDi("contexts")(1)("context")
        
        Set paramsBiDi = New Dictionary
        paramsBiDi.Add "context", targetContext
        Set resultBiDi = bidiPlus.invokeMethod("goog:cdp.getSession", paramsBiDi)
        
        If Not resultBiDi Is Nothing Then
             MsgBox "現在のタブ(Context)に紐づく、裏側の『CDPセッションID』を取得しました！" & vbCrLf & vbCrLf & _
                    "【SessionID】" & resultBiDi("session"), vbInformation, "BiDi+ GetSession"
                    
             Dim cdpSessionId As String
             cdpSessionId = resultBiDi("session")
        End If
    End If

    '-----------------------------------------------------------------------
    ' 2. goog:cdp.sendCommand を使って、CDPの「Browser.getVersion」を実行してみる
    '-----------------------------------------------------------------------
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "method", "Browser.getVersion"
    paramsBiDi.Add "params", New Dictionary
    If cdpSessionId <> "" Then paramsBiDi.Add "session", cdpSessionId
    
    Set resultBiDi = bidiPlus.invokeMethod("goog:cdp.sendCommand", paramsBiDi)
    
    If Not resultBiDi Is Nothing Then
        MsgBox "CDPコマンド(Browser.getVersion)をBiDi経由で実行できました！" & vbCrLf & vbCrLf & _
               "【Browser】" & resultBiDi("result")("userAgent") & vbCrLf & _
               "【Protocol-Version】" & resultBiDi("result")("protocolVersion"), vbInformation, "BiDi+ CDP Tunnel"
    End If

    '終了
    bidiPlus.quit
End Sub



'***************************************************************************************************
'                               ■■■ リアタッチDemo ■■■
'***************************************************************************************************
'* 機能　　：複数プロシージャをまたがった段階的な処理を行う際の再接続Demoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：単一プロシージャで完結出来ない場面がきっとあるはずです。途中でセキュリティ認証による手作業が入ったりなど...
'            そういった場面でも、デバックブラウザで起動済みへ再接続するDemoです
'***************************************************************************************************
Sub demoReattachmentPart1()
    ' 起動
    Dim First As WebDriverBiDiCore
    Set First = 設定シートからのBiDi起動

    '現在のコンテキストIDを取得する
    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary
    Set resultBiDi = First.invokeMethod("browsingContext.getTree")
    Dim targetContext As String
    If Not (resultBiDi Is Nothing) Then
        targetContext = resultBiDi("contexts")(1)("context")    '一旦は、先頭タブで　※本来はURLcheckとかがいると思うが、低レベル制御の都合上、妥協
    End If

    'GoogleTopページへ遷移
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "context", targetContext
    paramsBiDi.Add "url", "https://www.google.com/"
    paramsBiDi.Add "wait", "complete"
    First.invokeMethod "browsingContext.navigate", paramsBiDi
End Sub

Sub demoReattachmentPart2()
    '設定セルから、ユーザ名を取得
    With ShSetting01_StartBrowser
        Dim UserName As String
        UserName = .Range(.UseRangeName(2, "Demo_CDP.demoReattachmentPart2")).value
    End With

    ' リアタッチとして起動
    Dim Reattachment As New WebDriverBiDiCore
    Dim ResultReattach As Boolean
    ResultReattach = Reattachment.reattach(UserName)

    If Not (ResultReattach) Then Debug.Print "Failed to reattach. `demoReattachmentPart1`を始動しましたか？": Exit Sub

    '現在のコンテキストIDを取得する
    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary
    Set resultBiDi = Reattachment.invokeMethod("browsingContext.getTree")
    Dim targetContext As String
    If Not (resultBiDi Is Nothing) Then
        targetContext = resultBiDi("contexts")(1)("context")    '一旦は、先頭タブで　※本来はURLcheckとかがいると思うが、低レベル制御の都合上、妥協
    End If

    'wikipediaページへ遷移
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "context", targetContext
    paramsBiDi.Add "url", "https://wikipedia.com"
    paramsBiDi.Add "wait", "complete"
    Reattachment.invokeMethod "browsingContext.navigate", paramsBiDi
End Sub



'***************************************************************************************************
'                               ■■■ アップデートDemo ■■■
'***************************************************************************************************
'* 機能　　：ChromiumをBiDi制御する際の核となる`mapperTab.js`の更新Demoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ローカルファイル(オフライン) or NPM(jsdelivr-オンライン)経由による2パターンを提供します
'***************************************************************************************************
Private Sub ローカルファイルで更新()
    '1. ファイルパスを、ダイアログで指定
    Dim UpdateFilePath As String
    UpdateFilePath = Application.GetOpenFilename("mapperTab File, *.js", , "WebDriverBiDiのやりとりの基となる`mapperTab.js`相当を選択してください")
    If UpdateFilePath = CStr(False) Then Exit Sub

    '2. 一旦、パスごとに分割
    Dim tmp: tmp = Split(UpdateFilePath, "\")
    Dim SplitNum As Long: SplitNum = UBound(tmp)

    '3. 分けて格納
    Dim FolderName As String, FileName As String
    FileName = tmp(SplitNum): tmp(SplitNum) = ""
    FolderName = Join(tmp, "\")

    '4. 更新処理
    Dim UpdateBiDi As New WebDriverBiDiCore
    If (UpdateBiDi.UpdateFromLocalFile(FolderName, FileName)) Then MsgBox "アップデートに成功しました。" & vbCrLf & UpdateFilePath, vbInformation, "Success" Else MsgBox "アップデートに失敗しました。" & vbCrLf & UpdateFilePath, vbCritical, "failure"

    '5. ローカルファイルで更新した旨を記録
    With ShLibrary01_JS
        .Range(.UseRangeName(1, "Demo_WebDriverBiDi.npm経由で更新")).value = " Local"
    End With
End Sub

Private Sub npm経由で更新()
    Dim UpdateBiDi As New WebDriverBiDiCore
    With ShLibrary01_JS
        '1. 現在のバージョン確認
        Dim mapperTab_npmVersion        As String: mapperTab_npmVersion = UpdateBiDi.UpdateCheckNPMVersion
        Dim mapperTab_WorkSheetVersion  As Range: Set mapperTab_WorkSheetVersion = .Range(.UseRangeName(1, "Demo_WebDriverBiDi.npm経由で更新"))
        If mapperTab_npmVersion = mapperTab_WorkSheetVersion.value Then MsgBox "すでに`mapperTab.js`は、最新バージョンです。", vbExclamation, "既に最新です(" & mapperTab_WorkSheetVersion & ")": Exit Sub

        '2. npmで更新
        If UpdateBiDi.UpdateFromNPMFile Then MsgBox "npm経由で、アップデートに成功しました。", vbInformation, "Success(" & mapperTab_WorkSheetVersion & " → " & mapperTab_npmVersion & ")" Else MsgBox "npm経由での、アップデートに失敗しました。", vbCritical, "failure": Exit Sub

        '3. バージョンをワークシートに記録
        mapperTab_WorkSheetVersion.value = mapperTab_npmVersion
    End With
End Sub
