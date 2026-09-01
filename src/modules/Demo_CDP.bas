Attribute VB_Name = "Demo_CDP"
'===================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'---------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       22/01/26 Long Vh: added demoMultiProfileOperation and demoReattachment examples
'       07/01/26 Long Vh: update the sub procedures to show case the new .notify function
'       27/04/23 Long Vh: made many improvements with v2.5 to make methods even more intuitive.
'       07/06/22 Long Vh: corrected typos in comments + more examples
'       03/06/22 Long Vh: codes edited + notes added + added extensive comments for HSBC colleagues
' References  :
'       Microsoft Scripting Runtime
' Notes       :
'       The framework does not need a matching webdriver as this is not a webdriver-based API.
'       This module includes a few examples of automating browsers using CDP. For the
'       engine codes, refer to the class modules CDPBrowser, CDPCore, CDPElement, and WebJsonConverter
'       For original examples, refer to Chris' article on CodeProject:
'       https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'       For the latest update of the CDP Framework by Long Vh:
'       https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA
'===================================================================================================
Option Explicit
Option Private Module



'***************************************************************************************************
'                                  ■■■ 全ての始まり ■■■
'***************************************************************************************************
Sub CDPによる冒険の始まり()
    '設定シートに基づくブラウザ立ち上げ
    Dim HelloWorldAutomationBrowser As CDPContext
    Set HelloWorldAutomationBrowser = ShSetting01_StartBrowser.StartCDPModeContext

    '↓ここから、あなたのイメージをコードに落とし込む↓




    'ブラウザを正常に閉じる
    HelloWorldAutomationBrowser.ThisCDPBrowser.quit
End Sub



'***************************************************************************************************
'                               ■■■ Demoプロシージャ ■■■
'***************************************************************************************************
'* 機能　　：イベントキャプチャに関するDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：例えば、認証用URLのNetwork.loadingFinished を検知したら、そこの requestId から `Network.getResponseBody` を実行しToken入手なんてことが可能です。(でも、Token抽出とかはNetwork.getCookies や DOMStorage.getDOMStorageItems 等が楽です。)
'* 注意事項：ここでは、ネットワークイベントのデモですが、他のイベントも同じ操作でとらえることができます
'***************************************************************************************************
Sub ネットワークイベントの確認()
    '必要な変換オブジェクトを用意
    Dim CharConvObj As New CharacterCodeConversion

    '設定シートに基づくブラウザ立ち上げ
    Dim Demo_NetworkEvent As CDPContext: Set Demo_NetworkEvent = ShSetting01_StartBrowser.StartCDPModeContext

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
    Demo_NetworkEvent.ThisCDPBrowser.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "Event.json"


    '-------------------------------- 機能2：セーブデータを作成し、イベントキャプチャを無効化する --------------------------------
    Dim SaveDataEvents As Dictionary: Set SaveDataEvents = Demo_NetworkEvent.BrowserEvents  'セーブデータ作成
    Set Demo_NetworkEvent.BrowserEvents = Nothing               '`Nothing`を渡すことで、イベントを破棄するようになる


    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/youtube/20200714b.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.ThisCDPBrowser.TakeEvents

    'イベント情報をDownloadsフォルダに保存しますが、無効中なので0バイトになります
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "NotEvent.json"


    '-------------------------------- 機能3：セーブデータを読み込み、そこからイベントキャプチャを再開する --------------------------------
    Set Demo_NetworkEvent.BrowserEvents = SaveDataEvents        '既存のセーブデータを読み込む


    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/index.stm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.ThisCDPBrowser.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(WebJsonConverter.serialize(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "EventFromSaveData.json"


    'ブラウザを閉じる。demo終了
    Demo_NetworkEvent.ThisCDPBrowser.quit
End Sub

'***************************************************************************************************
'* 機能　　：日本語に関するDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：id属性やname属性に日本語が使われてるサイトでの動作テストです。コードは、`https://qiita.com/yaju/items/0807cc762af4a0568806`を参考にしてます。
'* 注意事項：このテストを行う際は、シート：ブラウザ起動設定 にて、`常にUTF-8でCDP-Json送信`をONにしてください
'***************************************************************************************************
Sub JapaneseElementTest()
    '設定シートに基づくブラウザ立ち上げ、体脂肪率計算サイトへアクセスします
    Dim Demo_Japanese As CDPContext: Set Demo_Japanese = ShSetting01_StartBrowser.StartCDPModeContext("https://keisan.site/exec/system/1161228728")

    ' 身長をセット
    Dim height As CDPElement
    Set height = Demo_Japanese.getElementByID("var_身長")

    '日本語と絵文字入力テスト
    height.sendString "うみねこ！" & WorksheetFunction.Unichar(128566) & WorksheetFunction.Unichar(8205) & WorksheetFunction.Unichar(127787) & WorksheetFunction.Unichar(65039) & "みゃ～お！" & WorksheetFunction.Unichar(129442)  '日本語兼サロゲートペア絵文字入力テスト(U+1F636 U+200D U+1F32B U+FE0F、U+1F9A2)
    Demo_Japanese.notify "身長を入力しました" & WorksheetFunction.Unichar(129418)       '日本語兼絵文字通知表示テスト(U+1F98A)
    CDPHelpers.Sleep 3

    'ちゃんと数字で入力しなおす
    height.sendString "170.5"
    Demo_Japanese.notify "身長を入力し直しました" & WorksheetFunction.Unichar(128397) & WorksheetFunction.Unichar(65039)    '日本語兼サロゲートペア絵文字通知表示テスト(U+1F58D U+FE0F)
    CDPHelpers.Sleep 3

    ' 体重をセット
    Dim weight As CDPElement
    Set weight = Demo_Japanese.getElementByID("var_体重")
    weight.sendString "48.5"
    Demo_Japanese.notify "体重を入力しました" & WorksheetFunction.Unichar(9878) & WorksheetFunction.Unichar(65039)    '日本語兼サロゲートペア絵文字通知表示テスト(U+2696 U+FE0F)
    CDPHelpers.Sleep 3

    ' ボタンクリック
    Demo_Japanese.getElementByID("executebtn").SimpleClick
    Demo_Japanese.wait
    Demo_Japanese.notify "体脂肪率を計算しました" & WorksheetFunction.Unichar(129518)    '日本語兼絵文字通知表示テスト(U+1F9EE)
    CDPHelpers.Sleep 3

    ' 体脂肪率を取得
    Dim 体脂肪率 As Double
    体脂肪率 = Demo_Japanese.getElementByID("ans1").innerText
    Debug.Print "体脂肪率は、" & 体脂肪率 & "% です。"


    'ブラウザを閉じる。demo終了
    Demo_Japanese.ThisCDPBrowser.quit
End Sub

'***************************************************************************************************
'* 機能　　：拡張機能を読み込むDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ブラウザ自身をターゲットとした`ExecuteCDP`の使用例です
'* 注意事項：・このテストを行う際は、シート：ブラウザ起動設定 にて、`CDP-Jsonで拡張機能を制御`をONにしてください
'            ・`Extensions`は実験的ドメインですが、Class内Err.Raiseでは止めずに、ここの自力判定でエラーハンドリングします
'***************************************************************************************************
Sub UseExtensions()
    '拡張機能があるアンパックフォルダパスを、ダイアログで指定
    '参考 → https://qiita.com/studio_haneya/items/9f5141b667efc3bfa615
    Dim ExtensionsFolderPath As String
    With Application.FileDialog(4)  'msoFileDialogFolderPicker
        .Title = "拡張機能の基となる`manifest.json`を含むフォルダを選択してください"
        .InitialFileName = Environ("UserProfile") & "\AppData\Local"    '初期位置

        If .show = -1 Then ExtensionsFolderPath = .SelectedItems(1) Else Exit Sub
    End With


    '設定シートに基づくブラウザ立ち上げ
    Dim controlExtensions As CDPContext: Set controlExtensions = ShSetting01_StartBrowser.StartCDPModeContext

    '拡張機能のページへ遷移
    controlExtensions.navigate "edge://extensions/"

    '拡張機能を読み込む
    Dim CDPparams As Dictionary, ResultCDP As BiDiCDPJson
    Set CDPparams = New Dictionary
    CDPparams.Add "path", ExtensionsFolderPath
    Set ResultCDP = controlExtensions.ThisCDPBrowser.ExecuteCDP("Extensions.loadUnpacked", CDPparams, False)    '今回は、エラー無視で設定

    '読み込まれたか確認する
    '※コマンド実行に失敗すると、`nothing`で返るので、この仕様を利用します
    If ResultCDP Is Nothing Then
        'CDP-Json結果に`error`要素あり
        MsgBox "拡張機能のインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.ThisCDPBrowser.LastCDPJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.ThisCDPBrowser.LastCDPJsonError("code")

        'ブラウザを閉じる。demo終了
        controlExtensions.ThisCDPBrowser.quit
        Exit Sub

    ElseIf ResultCDP.ExistsKey("id") Then
        MsgBox "拡張機能のインストールに成功しました。ブラウザをご確認ください。" & vbCrLf & "なお、OKを押すと、アンインストールします。", vbInformation, "ExtensionsID：" & ResultCDP("id")

    Else
        MsgBox "インストールIDの確認が取れませんでした。" & vbCrLf & vbCrLf & "<RawResult>" & vbCrLf & ResultCDP.Stringify, vbExclamation, "Not found id"

        'ブラウザを閉じる。demo終了
        controlExtensions.ThisCDPBrowser.quit
    End If


    '拡張機能をアンインストール
    Set CDPparams = New Dictionary
    CDPparams.Add "id", ResultCDP("id")
    Set ResultCDP = controlExtensions.ThisCDPBrowser.ExecuteCDP("Extensions.uninstall", CDPparams, False)

    '消えたか確認する
    If ResultCDP Is Nothing Then
        'CDP-Json結果に`error`要素あり
        MsgBox "拡張機能のアンインストールに失敗しました。" & vbCrLf & vbCrLf & "＜原因＞" & vbCrLf & controlExtensions.ThisCDPBrowser.LastCDPJsonError("message"), vbCritical, "ErrorCode:" & controlExtensions.ThisCDPBrowser.LastCDPJsonError("code")

    Else
        MsgBox "拡張機能のアンインストールに成功しました。ブラウザをご確認ください。", vbInformation, "Uninstall Done!"
    End If


    'ブラウザを閉じる。demo終了
    controlExtensions.ThisCDPBrowser.quit
End Sub

'***************************************************************************************************
'* 機能　　：JavaScript関数、`alert`処理に関するDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：非同期実行、イベントキャプチャした内容をもとにコマンド実行といったことをデモンストレーションします
'* 注意事項：ここでの非同期clickは`jsEval`で表現します
'***************************************************************************************************
Sub TestAlert()
    '設定シートに基づくブラウザ立ち上げ。`selenium`の独自テストページに遷移します
    Dim Demo_alerts As CDPContext: Set Demo_alerts = ShSetting01_StartBrowser.StartCDPModeContext("https://www.selenium.dev/selenium/web/alerts.html")


    '必要な変数を用意
    Dim paramsCDP As New Dictionary
    Dim resCDP As BiDiCDPJson
    Dim searchId As String
    Dim nodeId As Long
    Dim x As Double, y As Double

    'テキスト入力用のAlertに入力させる文字列の指定
    Dim 入力文字内容 As String: 入力文字内容 = "VBAから入力したテスト文字列です！" & WorksheetFunction.Unichar(129418)


    With Demo_alerts
        ' --- 1. 必要なドメインを有効化 ---
        .ExecuteCDP ("Page.enable")

        Dim i As Long
        For i = 1 To 3
            Dim TargetXpath As String
            Select Case i
                Case 1: TargetXpath = "alert"
                Case 2: TargetXpath = "empty-alert"
                Case 3: TargetXpath = "prompt"
            End Select


            ' --- 6. 非同期でコマンド実行(Jsのクリック処理) ---
            'この瞬間、JavaScriptの`alert`関数が発動されます
            '※非同期処理を行うため、`CDPElement.cls`を使わない形で取ります
            Dim AsyncID As Long
            AsyncID = .jsEval("document.getElementById('" & TargetXpath & "').click()", RunAsyncCDP:=True)


            ' --- 7. イベントキャプチャを有効化 ---
            Set .BrowserEvents = New Dictionary


            ' --- 8. 特定のイベント名が出るまでループ ---
            Const SearchEventName As String = "Page.javascriptDialogOpening"
            Do
                '非同期イベントを取り出す
                .ThisCDPBrowser.TakeEvents

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
            .ExecuteCDP "Page.handleJavaScriptDialog", paramsCDP


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
        .ThisCDPBrowser.quit
    End With
End Sub

'***************************************************************************************************
'* 機能　　：WebView2を起動します
'---------------------------------------------------------------------------------------------------
'* 注意事項：`ICoreWebView2Settings`等の一部設定は、ページ遷移前のみ有効です
'***************************************************************************************************
Sub ExcelのユーザーフォームにWebView2を埋め込む()
    '1. UserForm側のWebView2の初期化を済ませる
    With WebView2Form
        If Not .StartCDPModeWebView2(addArgs:=ShSetting01_StartBrowser.UseRangeID(3, "Demo_CDP.ExcelのユーザーフォームにWebView2を埋め込む")) Then Debug.Print "WebView2の初期化に失敗しました。WebView2Loader.dllが見つからない、" & _
                                                        "またはEnvironment/Controllerの生成に失敗した可能性があります。": Exit Sub

        '2. 事前設定を施す(任意)
        .ThisWebView2.DevToolsEnabled = False
        .ThisWebView2.ContextMenuEnabled = False

        '3. ページ遷移
        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"

        '4. フォームを表示
        .show
    End With
End Sub

'***************************************************************************************************
'* 機能　　：ShadowRootに関するDemoです。シンプル版です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Open/Close 問わず、メソッドチェーン操作で利用できます
'* 注意事項：ここでの非同期clickは`CDPElement`で表現します
'***************************************************************************************************
Sub SimpleShadowRootTest()
    '1. ShadowRootページを開く
    Dim ShadowRootTest As CDPContext: Set ShadowRootTest = ShSetting01_StartBrowser.StartCDPModeContext("https://jec.fish/demo/shadow-open-close")
    With ShadowRootTest
        '2. Shadow-Root(Open) 内のボタンをクリック
        .getElementByXPath("//*[@id='open']/open-dom").GetShadowRoot.getElementByQuery("div > button").click

        '3. Shadow-Root(Close) 内のボタン要素を取得(`.click`にしない理由は以降のコメントに)
        Dim JavaScriptAlertButton As CDPElement
        Set JavaScriptAlertButton = .getElementByXPath("//*[@id='closed']/closed-dom").GetShadowRoot.getElementByQuery("div > button")

        '4. 次の操作前に下準備
        .pageEnable                           '`Page`ドメインを有効
        Set .BrowserEvents = New Dictionary   'イベントキャプチャを有効化

        '5. ボタン押下後、JavaScriptアラートが発動するため非同期実行するように設定(先述にて、直で`.click`をしないのはこのため)
        JavaScriptAlertButton.SetOptionRunAsyncCDP = True

        '改めてクリック処理
        JavaScriptAlertButton.SimpleClick
        JavaScriptAlertButton.SetOptionRunAsyncCDP = False  '元に戻しておく

        ' --- 6. 特定のイベント名が出るまでループ ---
        Const SearchEventName As String = "Page.javascriptDialogOpening"    'JavaScriptアラートが出るのでその検知
        Do
            '非同期イベントを取り出す
            .ThisCDPBrowser.TakeEvents

            'イベント名の確認
            If .BrowserEvents("EventMethods").Exists(SearchEventName) Then
                '出ているダイアログの情報の確認
                Dim tmp
                For Each tmp In .BrowserEvents("EventMethods")(SearchEventName)
                    Debug.Print "url    :"; tmp("params")("url")
                    Debug.Print "message:"; tmp("params")("message")
                    Debug.Print "type   :"; tmp("params")("type") & vbCrLf
                Next

                '1件、見つかったので少し待って、抜ける
                CDPHelpers.Sleep 2
                Exit Do
            End If
        Loop While True

        ' --- 7. ダイアログに反応しておく ---
        Dim paramsCDP As New Dictionary
        paramsCDP.Add "accept", True
        .ExecuteCDP "Page.handleJavaScriptDialog", paramsCDP
        CDPHelpers.Sleep

        '8. ブラウザを正常に閉じる
        .ThisCDPBrowser.quit
    End With
End Sub

'***************************************************************************************************
'* 機能　　：ShadowRootに関するDemoです。iframe内のShadowRoot操作編です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：iframe自体の`Target.getTargets`を取得して、そこのShadowRoot操作を行うイメージです
'***************************************************************************************************
Sub iframeShadowRootTest()
    '1. captchaDemoページを開く
    Dim captchaDemo As CDPBrowser: Set captchaDemo = ShSetting01_StartBrowser.StartCDPMode("https://2captcha.com/demo/cloudflare-turnstile")
    With captchaDemo
        '2. cloudflare 用のiframeにアタッチする
        Dim CloudflareTurnstile As CDPContext
        Set CloudflareTurnstile = .getTab(Url:="https://challenges.cloudflare.com/cdn-cgi/challenge-platform/", SearchTypeID:=kFrame, doRetrySecond:=5)     '※見つかるまで、5秒間内部でループされます

        '3. そのiframe内にあるチェックBoxをクリックする
        CloudflareTurnstile.getElementByQuery("body").GetShadowRoots(1).getElementByQuery("input").click    '本当は1個しかないですが、ここのDemoではあえて、複数用メソッドを使用します

        '4. 少し待って、閉じる
        CDPHelpers.Sleep 2
        .quit
    End With
End Sub

Sub RunChromium()
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
    Dim edge As CDPContext
    Set edge = ShSetting01_StartBrowser.StartCDPModeContext

   'Navigate and wait
   'If till argument is omitted, will by default wait until ReadyState = complete
    edge.navigate "https://livingwaters.com/movie/the-atheist-delusion/", isInteractive

   'Get view count via the new notify method
    Dim viewCount As Long
    viewCount = edge.getElementByQuery("[data-id='4b9a4b19']").innerText
    edge.notify "This free movie has already reached " & viewCount & " views! Wow!"

End Sub

Sub runHidden()
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
'   3. Click on the first result to reach the CodeProject's article
'   4. The vote count is seen there.
'
' ※日本国では、正しく機能しません。恐らく、検索地域の問題と思われます。
'---------------------------------------------------------------------------------

    Dim chrome As CDPContext

   'Start and hide
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext
    chrome.hide

   'Perform automation in the background
    chrome.navigate "https://google.com", isInteractive
    chrome.getElementByQuery("[name='q']").value = "automate edge vba"
    chrome.getElementByQuery("[name='q']").submit
    chrome.wait

   'Click the target result link
    chrome.getElementByXPath("//h3[text()='Automate Chrome / Edge using VBA']").click

   'Get the vote count only once the target element appears on screen
   'The onExists method is needed as this element appears after ReadyState = "complete"
    Dim voteCount As Long
    voteCount = chrome.getElementByID("ctl00_RateArticle_VoteCountNoHist").onExist.innerHTML

   'Confirm result and display
    Dim userChoice
    userChoice = MsgBox("Automation completed. Current vote counts: " & voteCount & ". Do you want to see the window?", vbYesNo)
    If userChoice = vbYes Then chrome.show Else chrome.ThisCDPBrowser.quit

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

    Dim chrome As CDPContext

    'Start and hide
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext
    chrome.hide

    'Perform automation in the background
    chrome.navigate "https://google.com", isInteractive
    chrome.getElementByQuery("[name='q']").value = "automate edge vba"
    chrome.getElementByQuery("[name='q']").submit
    chrome.wait '検索ボタン押下によりページ遷移発生につき、`wait`を挟む

    'Click the target result link
    chrome.getElementByXPath("//h3[text()='Chrome DevTools ProtocolでEdgeを操作するVBAマクロ']").click      '2026/02/16 時点での、最上位結果

    'Confirm result and display
    Dim userChoice As Long
    userChoice = MsgBox("Automation completed. Do you want to see the window?", vbYesNo)
    If userChoice = vbYes Then chrome.show Else chrome.ThisCDPBrowser.quit

End Sub

Sub runTabsAsOne()
'--------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' Similar to the runInstances example but this is with multiple tabs in
' the same instance instead.
'--------------------------------------------------------------------------

    Dim chrome As CDPContext
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext
    chrome.show

   'Automate Tabs
    chrome.Url = "https://google.com"   'or [chrome.navigate "https://google.com"]
    chrome.ThisCDPBrowser.newTab "https://sg.yahoo.com"
    chrome.ThisCDPBrowser.newTab "https://bing.com"

   'Resize to complete
    CDPHelpers.Sleep    'ちょこっとクールタイムが必要みたい
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

    Dim chrome As New CDPBrowser
    Set chrome = ShSetting01_StartBrowser.StartCDPMode

   'Create and assign tabs
    Dim tab1 As CDPContext
    Dim tab2 As CDPContext
    Dim tab3 As CDPContext
    Set tab1 = chrome.getTab(setMain:=True)     'The first tab is open by default after .start
    Set tab2 = chrome.newTab(newWindow:=True)   'newWindow: open tab as a new window instead of a tab
    Set tab3 = chrome.newTab(newWindow:=True)

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
    Dim chrome As CDPContext
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext
    'chrome.start addArgs:="--disable-popup-blocking"    'The disable-popup-blocking argument is needed to allow opening link in a new tab
    chrome.show asMaximized

   'Perform standard google search
    chrome.navigate "https://google.com"
    chrome.getElementByQuery("[name='q']").value = "newstarget.com"
    chrome.getElementByQuery("[name='q']").submit
    chrome.wait '検索ボタン押下によりページ遷移発生につき、`wait`を挟む

   'Google search result returns links that open in the same tab window
   'For this demonstration, we need to make it open in a new tab window instead
    Dim targetElement As CDPElement
    Set targetElement = chrome.getElementByXPath(".//a[contains(@href, 'https://www.newstarget.com/')]")
    targetElement.setAttribute "target", "_blank"   'Modify the element attribute to open in a new tab instead
    targetElement.click                             'Click the link, a new tab will be spontaneously open

   'Use getTabNew to quickly refer to the next newly open tab
    Dim targetTab As New CDPContext
    Set targetTab = chrome.ThisCDPBrowser.getTab
    targetTab.wait

   'Feed the top news title for today
    Dim firstTitle As String
    firstTitle = targetTab.getElementByQuery("div[class='Headline']").innerText
    targetTab.notify "Top popular headline for the day is """ & firstTitle & """."

End Sub

Sub runIFrame()
'--------------------------------------------------------------------------
' This example demonstrates the CDP Framework v2.5 getIFrame technique for
' accessing iFrame element intuitively, an improvement over 1.0:
' 1. The use of App Mode via appUrl argument of the .start method.
' 2. The use of getIframe to easily access iFrame elements on the web page.
' 3. Working with a complex web design where nested iFrames are employed.
'--------------------------------------------------------------------------

    Dim demoUrl As String
    demoUrl = "https://www.w3schools.com/html/tryit.asp?filename=tryhtml_iframe_height_width"

    Dim chrome As New CDPContext
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext(demoUrl)

    Dim iFrame1 As CDPElement
    Dim iFrame2 As CDPElement
    Set iFrame1 = chrome.getElementByID("iframeResult").getIFrame
    Set iFrame2 = iFrame1.getElementByQuery("iframe[title='Iframe Example']").getIFrame

    Dim txt As String
    txt = iFrame2.getElementByQuery("h1").innerText
    chrome.notify "Retrieved text from the iFrame: """ & txt & """"

End Sub

Sub getSnapShot()
'--------------------------------------------------------------------------
' This example demonstrates the easy handling of capturing a screenshot of
' the current page under CDP method. The second argument of the snapPage
' method can be set to True to capture the entire page or to False (default)
' to capture only the current view section of the page.
'--------------------------------------------------------------------------

    Dim demoUrl As String
    demoUrl = "https://www.google.com/search?q=1sgd+to+vnd"

    Dim chrome As CDPContext
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext   'not App Mode as sometimes Chrome App Mode does not allow file downloading
    chrome.navigate demoUrl

   'Snap a portion of the page based on the element indicator
   'If the second argument is omitted, snapPage will snap the entire page
    chrome.snapPage Environ("UserProfile") & "\Downloads", "todaySGDvsVND.png" 'chrome.snapPage(fileName, True) to capture the entire page instead

    Dim FileName As String: FileName = Environ("UserProfile") & "\Downloads\todaySGDvsVND.png"
    chrome.notify "Screenshot captured under " & FileName

End Sub

Sub fillReactForm()
'-------------------------------------------------------------------------
' This example demonstrates the power of 2.6 on working natively
' with React form fields, which are notoriously complex to automate
' due to the fact that React form uses its own internal event handlings.
' The demo aims to:
' 1. Fill in the name field on the page.
' 2. Press submit.
' 3. If the field input is recognized by React, alert will tell its value.
' Updated: 07/01/26: .sendKeys has been replaced with .sendString
'
' ※残念ながら、404により検証不可
'-------------------------------------------------------------------------

    Dim demoUrl As String
    demoUrl = "https://cdpn.io/gaearon/fullpage/VmmPgp?anon=true&editors=0010&view="

    Dim chrome As CDPContext
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext
    chrome.navigate demoUrl

   'Get the target fields
    Dim ip As CDPElement
    Dim sb As CDPElement
    Set ip = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='text']")
    Set sb = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='submit']")

   'This traditional input method will fail as this is a React field
'    chrome.jsEval ip.varName & ".value = 'TEST1'"
'    chrome.jsEval ip.varName & ".dispatchEvent(new Event('input', { bubbles: true, simulated: true }))"
    sb.click 'you will not see "TEST1" in the alert result

   'This will succeed by using 2.6-enhanced .value property
    ip.value = "TEST2" '.value property is now overloaded with a smart React field detection & inputing
    sb.click

   'This will succeed as it mimicks sending raw keys but to a specific element
    ip.sendString "TEST3"
    sb.click

End Sub

Sub switchMain()
'---------------------------------------------------------------
' This example demonstrate the use of argument setMain to switch
' the main session tab to another tab so that future
' reattachment will hook this tab directly. This is useful if
' the main tab is supposed to be a tab open subsequently during
' the automation process by the target web link. The setMain
' method is preferrable to using "Set chrome = chrome.getTab..."
' because the latter method does not update the serial string
' for future reattachment.
'---------------------------------------------------------------

    Dim chrome As CDPContext
    Set chrome = ShSetting01_StartBrowser.StartCDPModeContext
    chrome.ThisCDPBrowser.newTab "http://google.com", setMain:=True  'the chrome object will now directly refer to the Google tab
    chrome.ThisCDPBrowser.getTab("about:blank").closeTab             'prior 2.7, the next line will throw an error due to no main-switching mechanism
    chrome.printParams

End Sub

Sub demoMultiProfileOperation()
'----------------------------------------------------------------------------------------
' This example demonstrates a powerful feature of v3.1 called multi-instances operation.
' Under multi-instances, our framework can open browsers as separate independent
' instances; thereby enables advanced automation tactics such as robotic process
' automation and asynchronous operation. The procedure below attempts to open 2 CDP
' instances and runs at the same time asynchronously - something that VBA natively does
' not support. You will be able to see from the Immediate Window that (1) execBot2 is
' started first then execBot1 and (2) both bot operations run simultaneously and bot 1
' finishes first (likely as yahoo.com has less thing to load then finance.yahoo.com)
' even though it is started after bot 2. This implies that thanks to the multi-
' instances framework, we can achieve asynchronous operation.
'
' Without this feature, the closest to this application is to open a CDP session with
' multiple tabs but in that scenario, you can not achieve asynchronous operation as
' Chrome Devtools Protocol is tied to a single user profile so automation on a tab has
' to wait for one another. Additionally, if one tab causes the browser to crash, other
' running tabs will crash as well.
'----------------------------------------------------------------------------------------

    Application.OnTime Now + TimeValue("00:00:01"), "execBot1"
    execBot2

End Sub

Function execBot1()
'----------------------------------------------------------------------------------------
' Refer to the demoMultiProfileOperation
'----------------------------------------------------------------------------------------

    Debug.Print Format(Now, "hh:mm:ss") & " execBot1 started."

    Dim e1 As CDPContext
    Set e1 = ShSetting01_StartBrowser.StartCDPModeContext
    e1.navigate "https://yahoo.com"

    Debug.Print Format(Now, "hh:mm:ss") & " execBot1 completed."

End Function

Function execBot2()
'----------------------------------------------------------------------------------------
' Refer to the demoMultiProfileOperation
'----------------------------------------------------------------------------------------

    Debug.Print Format(Now, "hh:mm:ss") & " execBot2 started."

    Dim e2 As CDPContext
    Set e2 = ShSetting01_StartBrowser.StartCDPModeContext(SwitchUser:="CDP2")
    e2.navigate "https://finance.yahoo.com"

    Debug.Print Format(Now, "hh:mm:ss") & " execBot2 completed."

End Function



'***************************************************************************************************
'                               ■■■ リアタッチDemo ■■■
'***************************************************************************************************
'* 機能　　：複数プロシージャをまたがった段階的な処理を行う際の再接続Demoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：単一プロシージャで完結出来ない場面がきっとあるはずです。途中でセキュリティ認証による手作業が入ったりなど...
'            そういった場面でも、デバックブラウザで起動済みへ再接続するDemoです
'***************************************************************************************************
Sub demoReattachmentPart1()

    Dim c As CDPContext
    Set c = ShSetting01_StartBrowser.StartCDPModeContext
    c.navigate "https://google.com"

'    c.KeepSession = True    'もし、SessionIDを保持する場合はこれを最後に足して`demoReattachmentPart2ForTab`にてお試しください
End Sub

'***************************************************************************************************
'* 機能　　：ブラウザのパイプハンドルの接続まで担うリアタッチです
'---------------------------------------------------------------------------------------------------
'* 注意事項：・あくまでも、ブラウザの接続までです。その後のContext(タブ)接続は、手動で`getTab` OR `newTab`で出来ます
'            ・ブラウザのパイプハンドルが生きてない場合は、VBAエラーになります。`demoReattachmentPart1`からやり直しです
'***************************************************************************************************
Sub demoReattachmentPart2ForBrowser()
    Dim c As New CDPBrowser
    Dim r As CDPContext

    '設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '1. Excelに記録されてるパイプハンドル情報から復旧を試みる
    c.reattachPipe UserName

    '2. 未接続のタブに接続
    '※この時、必ず`setMain:=True`とすること。必要に応じて検索条件(URLマッチ等)も設定して下さい
    Set r = c.getTab(setMain:=True)
'    Set r = c.newTab(setMain:=True) '新しいタブ生成からでもOK

    '3．別ページに遷移して終了
    r.navigate "https://kemono-friends.jp/"
End Sub

'***************************************************************************************************
'* 機能　　：Context(タブ)接続まで担うリアタッチです
'---------------------------------------------------------------------------------------------------
'* 注意事項：Context(タブ)情報が失ってる場合は、このDemoではエラーとなります
'***************************************************************************************************
Sub demoReattachmentPart2ForTab()
    Dim c As New CDPContext

    '設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '1. Excelに記録されてる`TargetID`の生存確認
    '※第2引数で、Excelに記録されてる`SessionId`の使いまわしの設定が可能です。事前に`KeepSession = True`と書く必要はあります。
    If Not c.reattachPipe(UserName, False) Then MsgBox "「" & UserName & "」に接続できませんでした。TargetID情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub

    '2．再接続できたので、別ページに遷移して終了
    c.navigate "https://kemono-friends-20170110.jp/"
End Sub



'***************************************************************************************************
'                               ■■■ WebSocket経由版Demo ■■■
'***************************************************************************************************
'* 機能　　：`--remote-debugging-port`や「edge://inspect/#remote-debugging」に接続する際の簡易Demoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：タブへ接続します
'* 注意事項：・`WebSocket`という「後付け」の特性上、接続を確立後、`reattach`に渡す方式をとってます
'            ・事前に、デバッグブラウザの起動を済ませる必要があります
'***************************************************************************************************
Sub AutoConnectTab()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. 指定のWebSocketForCDPへ接続
    Dim WebSocketCDP As New CDPCoreViaWebSocket
    Debug.Print WebSocketCDP.AutoConnectPageCDP(UserName)

    '3. 繋げたWebSocketオブジェクトを`reattachWebSocket`メソッドに渡す
    Dim t As New CDPContext
    If Not t.reattachWebSocket(UserName, WebSocketCDP) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub

    '4. ページ遷移
    t.navigate "https://www.youtube.com/@islandfox6864"

    '5. WebSocketから切断
    WebSocketCDP.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：ローカルブラウザ起動から一通りの制御を行います
'***************************************************************************************************
Sub AutoConnectBrowser()
    '1. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '2. WebSocket制御で、ブラウザを起動
    Dim WebSocketCDP As New CDPCoreViaWebSocket
    Dim BrowserControl As CDPBrowser
    Set BrowserControl = WebSocketCDP.RunWebSocketModeBrowserCDP(IIf(ShSetting01_StartBrowser.UseRangeID(4, "Demo_CDP.AutoConnectBrowser"), BrowserList.RunChrome, BrowserList.RunEdge), , UserName, ShSetting01_StartBrowser.UseRangeID(3, "Demo_CDP.AutoConnectBrowser"))

    '3. 未接続のタブに接続
    Dim t As CDPContext
    Set t = BrowserControl.getTab(setMain:=True)

    '5. ページ遷移
    t.navigate "https://www.youtube.com/@direwolf8958/"

    '6. WebSocketから切断
    WebSocketCDP.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：今、目の前のブラウザを制御します
'***************************************************************************************************
Sub AutoConnectDevToolsActivePort()
    '1. 指定のWebSocketForCDPへ接続
    '※「edge://inspect/#remote-debugging」にて事前準備が必要です
    Dim WebSocketCDP As New CDPCoreViaWebSocket
    Debug.Print WebSocketCDP.AutoConnectDevToolsActivePort

    '2. 繋げたWebSocketオブジェクトを`reattachWebSocket`メソッドに渡す
    Dim b As New CDPBrowser
    b.reattachWebSocket "User Data", WebSocketCDP

    '3. 未接続のタブに接続
    Dim t As CDPContext
    Set t = b.newTab(setMain:=True) '新しいタブ生成からでもOK

    '4. ページ遷移
    t.navigate "https://www.youtube.com/@large-spottedgenet4617/"

    '5. WebSocketから切断
    WebSocketCDP.DisconnectCDP
End Sub

'***************************************************************************************************
'* 機能　　：このExcelで起動中のWebView2を乗っ取って、新規タブからスクレイピング操作を開始します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：・Excelの一部の操作はWebView2が動いてます。この仕様を利用して、デバッグポートを開けてそこから制御を行います
'            ・ここからの制御の場合、「RemoteDebuggingAllowed」のポリシー規制をスルー出来るようです
'
'* 注意事項：・VBEからの起動では失敗します。ワークシート上にある図形に「マクロの登録」でこのプロシージャを登録して、その図形から起動しないと機能しません
'            ・裏技チックのため、いつか使えなくなるかもしれません
'            ・起動に失敗する場合は、該当のWebView2プロセスをKillして下さい
'            ・既存タブではURL遷移に制限があるため、新しいタブを生成しそこからスクレイピングを始めれば今まで通りのスクレイピングが可能です
'***************************************************************************************************
Sub OpenExcelWebView2()
    '1. デバッグ用のポートをOpen
    Dim HelpWebView2 As New CDPCoreViaWebSocket
    HelpWebView2.EnsureWebView2DebugPort = 9222

    '2. Helpを開いて、疑似的にWebView2を始動させる
    CommandBars.ExecuteMso "Help"

    '3. 設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.CurrentUserName

    '4. 指定のWebSocketForCDPへ接続
    Debug.Print HelpWebView2.AutoConnectBrowserCDP(UserName)

    '5. 繋げたWebSocketオブジェクトを`reattachWebSocket`メソッドに渡す
    Dim b As New CDPBrowser
    b.reattachWebSocket UserName, HelpWebView2

    '6. 新しいタブに接続
    Dim t As CDPContext
    Set t = b.newTab(setMain:=True)

    '7. ページ遷移
    t.navigate "https://www.youtube.com/@humboldtpenguin2619"

    '8. WebSocketから切断
    HelpWebView2.DisconnectCDP
    HelpWebView2.EnsureWebView2DebugPort = -1
End Sub
