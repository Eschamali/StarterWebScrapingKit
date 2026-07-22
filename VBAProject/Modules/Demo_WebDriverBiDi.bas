Attribute VB_Name = "Demo_WebDriverBiDi"
'==============================================================================================================
'               Automating Chromium-Based Browsers with WebDriverBiDi API and VBA
'==============================================================================================================
Option Explicit
Option Private Module



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
'* 注意事項：・Demoモジュールにあるコードですが、他の部分で共用してるため、消さずにどこかにカット&ペーストしておくとよいでしょう
'            ・現時点では、タブへの接続まで自動で行いません
'***************************************************************************************************
Public Function 設定シートからのBiDi起動(Optional StartURL As String, Optional SwitchUser As String, Optional KioskMode As edgeKioskType, Optional sessionCapabilitiesRequest As Dictionary) As WebDriverBiDiMode
    '設定シートの各セルから設定値を取得し、適用
    With ShSetting01_StartBrowser
        '起動ブラウザ種類の設定
        '※BiDi-Json コマンドによる操作ですが、Chromium系統に特化した制御のため、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim ブラウザ名 As String: ブラウザ名 = IIf(.UseRangeID(4, "Demo_WebDriverBiDi.設定シートからのBiDi起動"), "chrome", "edge")

        '第2引数が省略ならシート側の設定を適用
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .UseRangeID(2, "Demo_WebDriverBiDi.設定シートからのBiDi起動"), SwitchUser)

        'ブラウザ起動
        Set 設定シートからのBiDi起動 = New WebDriverBiDiMode
        設定シートからのBiDi起動.StartBiDiMode ブラウザ名, StartURL, UseDataDir, .UseRangeID(3, "Demo_WebDriverBiDi.設定シートからのBiDi起動"), KioskMode, sessionCapabilitiesRequest
    End With
End Function

Public Function 設定シートからのBiDi起動ForTab(Optional StartURL As String, Optional SwitchUser As String, Optional KioskMode As edgeKioskType, Optional sessionCapabilitiesRequest As Dictionary) As WebDriverBiDiContext
    '設定シートの各セルから設定値を取得し、適用
    With ShSetting01_StartBrowser
        '起動ブラウザ種類の設定
        '※BiDi-Json コマンドによる操作ですが、Chromium系統に特化した制御のため、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim ブラウザ名 As String: ブラウザ名 = IIf(.UseRangeID(4, "Demo_WebDriverBiDi.設定シートからのBiDi起動ForTab"), "chrome", "edge")

        '第2引数が省略ならシート側の設定を適用
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .UseRangeID(2, "Demo_WebDriverBiDi.設定シートからのBiDi起動ForTab"), SwitchUser)

        'ブラウザ起動
        Set 設定シートからのBiDi起動ForTab = New WebDriverBiDiContext
        設定シートからのBiDi起動ForTab.StartBiDiModeAndConnectTab ブラウザ名, StartURL, UseDataDir, .UseRangeID(3, "Demo_WebDriverBiDi.設定シートからのBiDi起動ForTab"), KioskMode, sessionCapabilitiesRequest
    End With
End Function

Sub BiDiによる冒険の始まり()
    '設定シートに基づくブラウザ立ち上げ
    Dim HelloWorldAutomationBrowser As WebDriverBiDiContext: Set HelloWorldAutomationBrowser = 設定シートからのBiDi起動ForTab

    '↓ここから、あなたのイメージをコードに落とし込む↓




    'ブラウザを正常に閉じる
    HelloWorldAutomationBrowser.InheritanceWebDriverBiDiMode.quit
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
    Dim CharConvObj As New CharacterCodeConversion

    'WebDriverBiDiの初期化とブラウザ立ち上げ
    Dim Demo_NetworkEvent As WebDriverBiDiContext
    Set Demo_NetworkEvent = 設定シートからのBiDi起動ForTab


    '-------------------------------- 機能1：イベントキャプチャを有効化する --------------------------------
    '`New Dictionary`を渡すことで、内部で非同期イベントの蓄積を開始する
    Dim resultBiDi As Dictionary
    Set Demo_NetworkEvent.InheritanceWebDriverBiDiMode.BiDiEvents = New Dictionary

    'BiDi側でネットワークイベントを購読開始する
    Dim paramsBiDi As Dictionary
    Set paramsBiDi = New Dictionary
    Dim eventsArray As New Collection
    eventsArray.Add "network.beforeRequestSent"
    eventsArray.Add "network.responseCompleted"
    eventsArray.Add "log.entryAdded"
    Set Demo_NetworkEvent.InheritanceWebDriverBiDiMode.sessionSubscribe = eventsArray

    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/excel/vba/file/file11.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う (念のため待機後にも余波を回収)
    Demo_NetworkEvent.InheritanceWebDriverBiDiMode.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.InheritanceWebDriverBiDiMode.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_Event.json"


    '-------------------------------- 機能2：セーブデータを作成し、イベントキャプチャを無効化する --------------------------------
    Dim SaveDataEvents As Dictionary: Set SaveDataEvents = Demo_NetworkEvent.InheritanceWebDriverBiDiMode.BiDiEvents  'セーブデータ作成
    Set Demo_NetworkEvent.InheritanceWebDriverBiDiMode.BiDiEvents = Nothing               '`Nothing`を渡すことで、イベント記録状態を破棄する


    'URL遷移
    Demo_NetworkEvent.navigate "http://officetanaka.net/youtube/20200714b.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.InheritanceWebDriverBiDiMode.TakeEvents

    'イベント情報をDownloadsフォルダに保存しますが、無効中なので破棄状態（0バイト等）になります
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.InheritanceWebDriverBiDiMode.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_NotEvent.json"


    '-------------------------------- 機能3：セーブデータを読み込み、そこからイベントキャプチャを再開する --------------------------------
    Set Demo_NetworkEvent.InheritanceWebDriverBiDiMode.BiDiEvents = SaveDataEvents        '既存のセーブデータを読み込む

    'URL遷移
    Demo_NetworkEvent.navigate "http://officetanaka.net/index.stm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.InheritanceWebDriverBiDiMode.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.InheritanceWebDriverBiDiMode.BiDiEvents)), Environ("UserProfile") & "\Downloads", "BiDi_EventFromSaveData.json"


    'ブラウザを閉じる。demo終了
    Demo_NetworkEvent.InheritanceWebDriverBiDiMode.quit
End Sub

'***************************************************************************************************
'* 機能　　：拡張機能を読み込むDemoコード(BiDi版)です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：BiDiプロトコルの `webExtension` モジュールを使用した拡張機能のインストール・アンインストールのデモです。
'* 注意事項：・このテストを行う際は、事前シート：ブラウザ起動設定 にて、`CDP-Jsonで拡張機能を制御` をONにしてください
'***************************************************************************************************
Sub UseExtensions()
    '拡張機能があるアンパックフォルダパスを、ダイアログで指定
    Dim ExtensionsFolderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "拡張機能の基となる`manifest.json`を含むフォルダを選択してください"
        .InitialFileName = Environ("UserProfile") & "\AppData\Local"    '初期位置

        If .show = -1 Then ExtensionsFolderPath = .SelectedItems(1) Else Exit Sub
    End With

    'WebDriverBiDiの初期化とブラウザ立ち上げ
    Dim controlExtensions As WebDriverBiDiContext

    '---- 拡張機能制御を有効化するオプションを作成 ---
    Dim caps As New Dictionary
    Dim alwaysMatch As New Dictionary

    ' BiDiでは、セッション確立時の引数として渡すか、WebDriver側のCapabilityで有効にする必要がありますが、
    ' CDPBrowserの仕組み（引数渡し）を利用するためそのまま起動します。
    caps.Add "capabilities", New Dictionary
    caps("capabilities").Add "alwaysMatch", alwaysMatch
    '-------------------------------------------------

    ' 起動
    Set controlExtensions = 設定シートからのBiDi起動ForTab(sessionCapabilitiesRequest:=caps)

    '拡張機能のテストページ（もしくは任意のページ）へ遷移
    controlExtensions.navigate "edge://extensions/"

    '-----------------------------------------------------------------------
    '拡張機能を読み込む (BiDi `webExtension.install`)
    '-----------------------------------------------------------------------
    Dim extData As New Dictionary, paramsBiDi As New Dictionary
    extData.Add "type", "path"
    extData.Add "path", ExtensionsFolderPath
    paramsBiDi.RemoveAll
    paramsBiDi.Add "extensionData", extData

    ' 今回はエラー無視で設定 (StopError:=False)
    Dim resultBiDi As BiDiCDPJson
    Set resultBiDi = controlExtensions.ExecuteBiDi("webExtension.install", paramsBiDi, False)

    '読み込まれたか確認する
    If resultBiDi Is Nothing Then
        ' コマンド実行に失敗した場合、LastBiDiJsonError からエラー情報を取得する
        MsgBox "拡張機能のインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.InheritanceWebDriverBiDiMode.LastBiDiJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.InheritanceWebDriverBiDiMode.LastBiDiJsonError("error")

        'ブラウザを閉じる。demo終了
        controlExtensions.InheritanceWebDriverBiDiMode.quit
        Exit Sub

    ElseIf resultBiDi.Exists("extension") Then
        ' BiDiの webExtension.install は `extension` キーで IDを返します。
        MsgBox "拡張機能のインストールに成功しました。ブラウザをご確認ください。" & vbCrLf & "なお、OKを押すと、アンインストールします。", vbInformation, "ExtensionsID：" & resultBiDi("extension")

    Else
        MsgBox "インストールIDの確認が取れませんでした。" & vbCrLf & vbCrLf & "<RawResult>" & vbCrLf & resultBiDi.Stringify, vbExclamation, "Not found id"

        'ブラウザを閉じる。demo終了
        controlExtensions.InheritanceWebDriverBiDiMode.quit
        Exit Sub
    End If

    '-----------------------------------------------------------------------
    '拡張機能をアンインストール (BiDi `webExtension.uninstall`)
    '-----------------------------------------------------------------------
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "extension", resultBiDi("extension")
    Set resultBiDi = controlExtensions.ExecuteBiDi("webExtension.uninstall", paramsBiDi, False)

    '消えたか確認する
    If resultBiDi Is Nothing Then
        MsgBox "拡張機能のアンインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.InheritanceWebDriverBiDiMode.LastBiDiJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.InheritanceWebDriverBiDiMode.LastBiDiJsonError("error")
    Else
        MsgBox "拡張機能のアンインストールに成功しました。ブラウザをご確認ください。", vbInformation, "Uninstall Done!"
    End If

    'ブラウザを閉じる。demo終了
    controlExtensions.InheritanceWebDriverBiDiMode.quit
End Sub

'***************************************************************************************************
'* 機能　　：JavaScript関数、`alert`処理に関するBiDi版のDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：非同期実行、イベントキャプチャした内容をもとにコマンド実行といったことをデモンストレーションします
'***************************************************************************************************
Sub TestAlert()
    'WebDriverBiDiの初期化とブラウザ立ち上げ
    Dim Demo_alerts As WebDriverBiDiContext

    '---- JavaScriptによる自動アラート処理を無効化するオプションを作成 ---
    Dim caps As New Dictionary

    Dim alwaysMatch As New Dictionary
    alwaysMatch.Add "unhandledPromptBehavior", "ignore"

    caps.Add "capabilities", New Dictionary
    caps("capabilities").Add "alwaysMatch", alwaysMatch
    '---------------------------------------------------------------------

    'オプションを適用させて、指定URLから直接起動
    Set Demo_alerts = 設定シートからのBiDi起動ForTab("https://www.selenium.dev/selenium/web/alerts.html", sessionCapabilitiesRequest:=caps)

    '結果とBiDiパラメーター変数を用意
    Dim paramsBiDi As Dictionary, resultBiDi As BiDiCDPJson

    'テスト入力文字列
    Dim 入力文字内容 As String: 入力文字内容 = "VBAから入力したテスト文字列です！" & WorksheetFunction.Unichar(129418)

    With Demo_alerts
        ' --- 1. 必要なドメイン(イベント)をサブスクライブ ---
        Dim eventsArray As New Collection
        eventsArray.Add "browsingContext.userPromptOpened"
        Set .InheritanceWebDriverBiDiMode.sessionSubscribe = eventsArray

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
            Set .InheritanceWebDriverBiDiMode.BiDiEvents = New Dictionary

            ' --- 3. 非同期でコマンド準備/実行(Jsのクリック処理) ---
            ' 対象の要素をクリックするJSを評価する
            Set paramsBiDi = New Dictionary
            paramsBiDi.Add "expression", "document.getElementById('" & targetID & "').click()"
            Dim targetDict As Dictionary
            Set targetDict = New Dictionary
            targetDict.Add "context", .context
            paramsBiDi.Add "target", targetDict
            paramsBiDi.Add "awaitPromise", False

            Dim AsyncID As Long
            'この瞬間、JavaScriptの`alert`関数が非同期で発動されます
            AsyncID = .InheritanceWebDriverBiDiMode.ExecuteBiDiAsync("script.evaluate", paramsBiDi)

            ' --- 4. 特定のイベント名が出るまでループ ---
            Const SearchEventName As String = "browsingContext.userPromptOpened"
            Do
                '非同期イベントを取り出す
                .InheritanceWebDriverBiDiMode.TakeEvents

                'イベント名の確認
                If .InheritanceWebDriverBiDiMode.BiDiEvents("EventMethods").Exists(SearchEventName) Then
                    '出ているダイアログの情報の確認
                    Dim tmp
                    For Each tmp In .InheritanceWebDriverBiDiMode.BiDiEvents("EventMethods")(SearchEventName)
                        Debug.Print "message:"; tmp("params")("message")
                        Debug.Print "type   :"; tmp("type") & vbCrLf
                    Next

                    '見つかったので抜ける
                    Exit Do
                End If
            Loop While True

            ' --- 5. ダイアログに反応しておく ---
            Set paramsBiDi = New Dictionary
            paramsBiDi.Add "accept", True
            paramsBiDi.Add "userText", 入力文字内容
            Set resultBiDi = .ExecuteBiDi("browsingContext.handleUserPrompt", paramsBiDi)

            ' --- 6. 以前、非同期で実行した結果も拝見する ---
'            Dim resBiDiAsync As Dictionary
'            .sleep 0.5 ' 結果取得のためのディレイ
'            .TakeEvents ' 受信キューを消化
'
'            Dim エラー確認 As Boolean
'            Set resBiDiAsync = .ResultBiDiForAsync(AsyncID, エラー確認)
'            If Not (resBiDiAsync Is Nothing) Then Debug.Print "resBiDiAsync - " & JsonDicObj.ConvertToJson(resBiDiAsync)

        Next

        ' --- 7. ブラウザを閉じる ---
        ' DOM経由のテキスト取得を、script.evaluateで代替
        Set paramsBiDi = New Dictionary
        paramsBiDi.Add "expression", "document.querySelector('#text > p') ? document.querySelector('#text > p').innerText : 'Not Found'"
        Set targetDict = New Dictionary
        targetDict.Add "context", .context
        paramsBiDi.Add "target", targetDict
        paramsBiDi.Add "awaitPromise", True
        Set resultBiDi = .ExecuteBiDi("script.evaluate", paramsBiDi)

        Dim Htmlの表示内容 As String
        If Not (resultBiDi Is Nothing) Then
            If resultBiDi.Exists("result") Then
                If resultBiDi("result").Exists("value") Then Htmlの表示内容 = resultBiDi("result")("value")
            End If
        End If

        Debug.Print "htmlの出力文字列：" & Htmlの表示内容
        Debug.Assert Htmlの表示内容 = 入力文字内容
        .InheritanceWebDriverBiDiMode.quit
    End With
End Sub

'***************************************************************************************************
'* 機能　　：BiDi+ (Chromium独自拡張) の `goog:cdp.sendCommand` を試すDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebDriver BiDi プロトコルにまだ存在しない詳細な機能を、従来のCDPコマンドを
'*           トンネリング（中継）して呼び出す「BiDi+」の機能デモンストレーションです。
'***************************************************************************************************
Sub TestBiDiPlus_CDPTunnel()
    Dim bidiPlus As WebDriverBiDiContext

    ' ブラウザ起動
    Set bidiPlus = 設定シートからのBiDi起動ForTab

    Dim paramsBiDi As Dictionary, resultBiDi As BiDiCDPJson

    '-----------------------------------------------------------------------
    ' 1. CDPのセッションIDを取得する (goog:cdp.getSession)
    '-----------------------------------------------------------------------
    Set paramsBiDi = New Dictionary
    Set resultBiDi = bidiPlus.ExecuteBiDi("goog:cdp.getSession", paramsBiDi)

    If Not resultBiDi Is Nothing Then
         MsgBox "現在のタブ(Context)に紐づく、裏側の『CDPセッションID』を取得しました！" & vbCrLf & vbCrLf & _
                "【SessionID】" & resultBiDi.StringKey("session"), vbInformation, "BiDi+ GetSession"

         Dim cdpSessionId As String
         cdpSessionId = resultBiDi("session")
    End If

    '-----------------------------------------------------------------------
    ' 2. goog:cdp.sendCommand を使って、CDPの「Browser.getVersion」を実行してみる
    '-----------------------------------------------------------------------
    Set paramsBiDi = New Dictionary
    paramsBiDi.Add "method", "Browser.getVersion"
    paramsBiDi.Add "params", New Dictionary
    If cdpSessionId <> "" Then paramsBiDi.Add "session", cdpSessionId

    Set resultBiDi = bidiPlus.ExecuteBiDi("goog:cdp.sendCommand", paramsBiDi)

    If Not resultBiDi Is Nothing Then
        MsgBox "CDPコマンド(Browser.getVersion)をBiDi経由で実行できました！" & vbCrLf & vbCrLf & _
               "【Browser】" & resultBiDi.NodeKey("result").StringKey("userAgent") & vbCrLf & _
               "【Protocol-Version】" & resultBiDi.NodeKey("result").StringKey("protocolVersion"), vbInformation, "BiDi+ CDP Tunnel"
    End If

    '終了
    bidiPlus.InheritanceWebDriverBiDiMode.quit
End Sub

Sub ConvertToCDPContextDemo()
    'WebDriverBiDiCoreの初期化とブラウザ立ち上げ
    Dim NewsSite As WebDriverBiDiMode
    Set NewsSite = 設定シートからのBiDi起動("https://news.google.com/home")

    'getタブでスマートにオブジェクト取得
    Dim BiDiTab As WebDriverBiDiContext
    Set BiDiTab = NewsSite.getTab("https://news.google.com/", setMain:=True)

    '別のURLへ遷移
    BiDiTab.navigate "https://m365.cloud.microsoft/chat"

    'CDP制御できるように変換
    Dim CDPTab As CDPContext
    Set CDPTab = BiDiTab.ConvertToCDPContext

    'CDP実行してみる
    CDPTab.notify "BiDiオブジェクトクラスから、CDP制御できるように変換できました！" & WorksheetFunction.Unichar(129418)
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
    Dim First As WebDriverBiDiContext
    Set First = 設定シートからのBiDi起動ForTab

    'GoogleTopページへ遷移
    First.navigate "https://www.google.com/"
End Sub

'***************************************************************************************************
'* 機能　　：WebDriverBiDi制御用タブの接続まで担うリアタッチです
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、WebDriverBiDi制御用のタブ接続までです。その後のContext(タブ)接続は、手動で`getTab` OR `newTab`で出来ます
'            ・ブラウザのパイプハンドルが生きてない場合は、エラーになります。`demoReattachmentPart1`からやり直しです
'            ・WebDriverBiDi制御用タブが無くなっても、`WebDriverBiDiMode`からの`reattach`で、再始動が可能です
'***************************************************************************************************
Sub demoReattachmentPart2()
    '設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_WebDriverBiDi.demoReattachmentPart2")

    '1. リアタッチとして起動
    Dim Reattachment As New WebDriverBiDiMode
    If Not Reattachment.reattach(UserName) Then Debug.Print "Failed to reattach. `demoReattachmentPart1`を始動しましたか？": Exit Sub

    '2. 未接続のタブに接続
    '※この時、必ず`setMain:=True`とすること。必要に応じて検索条件(URLマッチ等)も設定して下さい
    Dim ReattachmentTab As WebDriverBiDiContext
    Set ReattachmentTab = Reattachment.getTab(setMain:=True)
'    Set ReattachmentTab = Reattachment.newTab(setMain:=True)   '新しいタブ生成からでもOK

    '3. エラーチェック
    '※特に`getTab`の場合は、0個で返ることがあるのでそのチェックを行います
    If ReattachmentTab Is Nothing Then MsgBox "`browsingContext.getTree`の実行に成功しましたが、有効なタブが見つかりませんでした。" & vbCrLf & "ブラウザのタブを何個か開いてみて下さい。大抵は、2,3個程度追加で開けば、行けると思います。", vbCritical, "WebDriver BiDi": Exit Sub

    '4．別ページに遷移して終了
    ReattachmentTab.navigate "https://kemono-friends-20170110.jp/"
End Sub

'***************************************************************************************************
'* 機能　　：最後にWebDriverBiDiで制御したタブの接続まで担うリアタッチです
'---------------------------------------------------------------------------------------------------
'* 注意事項：最後にWebDriverBiDiで制御したタブが失ってる場合は失敗します
'***************************************************************************************************
Sub demoReattachmentPart2ForTab()
    '設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_WebDriverBiDi.demoReattachmentPart2ForTab")

    ' リアタッチとして起動
    Dim Reattachment As New WebDriverBiDiContext
    If Not Reattachment.reattach(UserName) Then MsgBox "「" & UserName & "」に接続できませんでした。`BiDi-context`情報がお亡くなりです。", vbCritical, "WebDriver BiDi": Exit Sub

    '別ページに遷移
    Reattachment.navigate "https://w3c.github.io/webdriver-bidi/"
End Sub



'***************************************************************************************************
'                               ■■■ WebSocket経由版Demo ■■■
'***************************************************************************************************
'* 機能　　：`--remote-debugging-port`や「edge://inspect/#remote-debugging」に接続する際の簡易Demoです
'---------------------------------------------------------------------------------------------------
'* 注意事項：・`WebSocket`という「後付け」の特性上、接続を確立後、`reattach`に渡す方式をとってます
'            ・事前に、デバッグブラウザの起動を済ませる必要があります
'            ・WebDriverBiDi制御用タブが無くなっても、`WebDriverBiDiMode`からの`reattach`で、再始動が可能です
'***************************************************************************************************
Sub SetupWebSocketMode()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_WebDriverBiDi.SetupWebSocketMode")

    '2. 指定のWebSocketForBiDiへ接続
    Dim WebSocketBiDi As New CDPCoreViaWebSocket
    Debug.Print WebSocketBiDi.AutoConnectBrowserCDP(UserName, True)         '基本はこっち。ExcelにあるBiDi制御タブ情報を流用するため、第2引数を`True`にしておく
'    Debug.Print WebSocketBiDi.AutoConnectDevToolsActivePort(UserName,True) '今、目の前のブラウザを制御する場合。ExcelにあるBiDi制御タブ情報を流用するため、第2引数を`True`にしておく

    '3. 繋げたWebSocketオブジェクトを`reattach`メソッドに渡す
    Dim m As New WebDriverBiDiMode
    If Not m.reattach(UserName, WebSocketMode:=WebSocketBiDi) Then Debug.Print "Failed to reattach. ブラウザの起動が必要です": Exit Sub

    '4. 新しいタブに接続
    Dim c As WebDriverBiDiContext
    Set c = m.newTab(setMain:=True)

    '5．別ページに遷移して終了
    c.navigate "https://www.youtube.com/@islandfox6864"

    '6. WebSocketから切断
    WebSocketBiDi.DisconnectCDP
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
    If UpdateBiDi.UpdateFromLocalFile(FolderName, FileName) Then MsgBox "ローカル`mapperTab.js`によるアップデートに成功しました。" & vbCrLf & UpdateFilePath, vbInformation, "Success" Else MsgBox "アップデートに失敗しました。Excelテーブルに埋め込んだJavaScript文字列とアップロードファイルとの一致が確認できませんでした。" & vbCrLf & UpdateFilePath, vbCritical, "failure"
End Sub

Private Sub npm経由で更新()
    Dim UpdateBiDi As New WebDriverBiDiCore
    With ShLibrary01_JS
        '1. 現在のバージョン確認
        Dim mapperTab_npmVersion    As String: mapperTab_npmVersion = UpdateBiDi.UpdateCheckNPMVersion
        Dim mapperTab_WorkSheetV    As String: mapperTab_WorkSheetV = ShLibrary01_JS.VersionMapperTabJS
        If mapperTab_npmVersion = mapperTab_WorkSheetV Then MsgBox "すでに`mapperTab.js`は、最新バージョンです。", vbExclamation, "既に最新です(" & mapperTab_WorkSheetV & ")": Exit Sub

        '2. npmで更新
        Dim UpdateSuccess As Boolean: UpdateSuccess = UpdateBiDi.UpdateFromNPMFile
        If UpdateSuccess Then MsgBox "npm経由で、アップデートに成功しました。", vbInformation, "Success(" & mapperTab_WorkSheetV & " → " & mapperTab_npmVersion & ")" Else MsgBox "npm経由での、アップデートに失敗しました。Excelに埋め込んだJavaScript文字列とnpm経由での一致が確認できませんでした。", vbCritical, "failure"

        '3. バージョンをワークシートに記録
        '※ミスの場合は、空欄にする
        ShLibrary01_JS.VersionMapperTabJS = IIf(UpdateSuccess, mapperTab_npmVersion, vbNullString)
    End With
End Sub
