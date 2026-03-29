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



'***************************************************************************************************
'                               ■■■ 設定プロシージャ ■■■
'***************************************************************************************************
'* 機能　　：設定シートから、パラメーターを読み込んで、CDPモードでブラウザを起動するヘルパープロシージャです
'---------------------------------------------------------------------------------------------------
'* 返り値　：クラスモジュール - CDPBrowser
'* 引数　　：StartURL   ブラウザ起動時にアクセスしたいURL。指定しない場合は、空ページ(abount:blank)になります。
'                       未指定でも クラスメソッド：navigate で後から、URL遷移も可能です。
'
'            SwtchUser  マルチインスタンス用に別ユーザーを指定するときに使用します
'            KioskMode  0(省略)：通常モード(キオスクモードは使いません)
'                       1      ：キオスクモード デジタル/対話型サイネージ
'                       2      ：キオスクモード パブリック ブラウジング
'---------------------------------------------------------------------------------------------------
'* 詳細説明：VBEによるハードコーディングではなく、設定シートから読み込む方式により、ユーザー側からも手軽に設定変更ができます
'* 注意事項：・Demoモジュールにあるコードですが、他の部分で共用してるため、消さずにどこかにカット&ペーストしておくとよいでしょう
'            ・Chromeにもキオスクモードはありますが、Edgeほど引数での調整はありません
'***************************************************************************************************
Public Function 設定シートからのCDP起動(Optional StartURL As String, Optional SwitchUser As String, Optional KioskMode As edgeKioskType) As CDPBrowser
    '設定シートの各セルから設定値を取得し、適用
    With ShSetting01_StartBrowser
        '起動ブラウザ種類の設定
        '※CDP－Json コマンドによる操作なので、Chromium系統であれば、Edge,Chrome 以外にもできるかと思いますが一旦はメジャーなやつのみで
        Dim ブラウザ名 As String: ブラウザ名 = IIf(.Range(.UseRangeName(4, "Demo_CDP.設定シートからのCDP起動")).value, "chrome", "edge")

        '第2引数が省略ならシート側の設定を適用
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .Range(.UseRangeName(2, "Demo_CDP.設定シートからのCDP起動")).value, SwitchUser)

        'ブラウザ起動
        Set 設定シートからのCDP起動 = New CDPBrowser
        設定シートからのCDP起動.start ブラウザ名, StartURL, .Range(.UseRangeName(6, "Demo_CDP.設定シートからのCDP起動")).value, UseDataDir, .Range(.UseRangeName(3, "Demo_CDP.設定シートからのCDP起動")).value, KioskMode
    End With
End Function

Sub CDPによる冒険の始まり()
    '設定シートに基づくブラウザ立ち上げ
    Dim HelloWorldAutomationBrowser As CDPBrowser: Set HelloWorldAutomationBrowser = 設定シートからのCDP起動

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
'* 注意事項：ここでは、ネットワークイベントのデモですが、他のイベントも同じ操作でとらえることができます
'***************************************************************************************************
Sub ネットワークイベントの確認()
    '必要な変換オブジェクトを用意
    Dim JsonDicObj As New WebJsonConverter
    Dim CharConvObj As New CharacterCodeConversion:
    
    '設定シートに基づくブラウザ立ち上げ
    Dim Demo_NetworkEvent As CDPBrowser: Set Demo_NetworkEvent = 設定シートからのCDP起動

    '一部の非同期イベントのみキャプチャするようにフィルターを設定
    '※未設定の場合は、全キャプチャとなります。今回のDemoの場合は、下記2つをコメントアウトすると、全キャプチャとなります
    Demo_NetworkEvent.SetFilterEvents = "Network.requestWillBeSent"
    Demo_NetworkEvent.SetFilterEvents = "Network.loadingFinished"


    '-------------------------------- 機能1：イベントキャプチャを有効化する --------------------------------
    Set Demo_NetworkEvent.BrowserEvents = New Dictionary        '`New Dictionary`を渡すことで、新規イベントキャプチャが可能になる。

    
    'ネットワークイベント受信を有効化する
    Dim ResultCDP As Dictionary: Set ResultCDP = Demo_NetworkEvent.invokeMethod("Network.enable")
    
    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/excel/vba/file/file11.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "Event.json"


    '-------------------------------- 機能2：セーブデータを作成し、イベントキャプチャを無効化する --------------------------------
    Dim SaveDataEvents As Dictionary: Set SaveDataEvents = Demo_NetworkEvent.BrowserEvents  'セーブデータ作成
    Set Demo_NetworkEvent.BrowserEvents = Nothing               '`Nothing`を渡すことで、イベントを破棄するようになる


    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/youtube/20200714b.htm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存しますが、無効中なので0バイトになります
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "NotEvent.json"


    '-------------------------------- 機能3：セーブデータを読み込み、そこからイベントキャプチャを再開する --------------------------------
    Set Demo_NetworkEvent.BrowserEvents = SaveDataEvents        '既存のセーブデータを読み込む
    

    'URL遷移して、読み込み終わるまで待機
    Demo_NetworkEvent.navigate "http://officetanaka.net/index.stm"

    '先ほどのURL遷移で発生した非同期イベントを取り出す処理を行う
    Demo_NetworkEvent.TakeEvents

    'イベント情報をDownloadsフォルダに保存
    CharConvObj.BytesToSaveFile CharConvObj.BytesFromString(JsonDicObj.ConvertToJson(Demo_NetworkEvent.BrowserEvents)), Environ("UserProfile") & "\Downloads", "EventFromSaveData.json"


    'ブラウザを閉じる。demo終了
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
    Dim Demo_Japanese As CDPBrowser: Set Demo_Japanese = 設定シートからのCDP起動("https://keisan.site/exec/system/1161228728")
    
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
    Demo_Japanese.quit
End Sub

'***************************************************************************************************
'* 機能　　：拡張機能を読み込むDemoコードです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ブラウザ自身をターゲットとした`invokeMethod`の使用例です
'* 注意事項：・このテストを行う際は、シート：ブラウザ起動設定 にて、`CDP-Jsonで拡張機能を制御`をONにしてください
'            ・`Extensions`は実験的ドメインですが、Class内Err.Raiseでは止めずに、ここの自力判定でエラーハンドリングします
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
    Dim controlExtensions As CDPBrowser: Set controlExtensions = 設定シートからのCDP起動
    
    '拡張機能のページへ遷移
    controlExtensions.navigate "edge://extensions/"

    '拡張機能を読み込む
    Dim CDPparams As Dictionary, ResultCDP As Dictionary
    Set CDPparams = New Dictionary
    CDPparams.Add "path", ExtensionsFolderPath
    Set ResultCDP = controlExtensions.invokeMethod("Extensions.loadUnpacked", CDPparams, True, False)   '今回は、エラー無視で設定

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
    Set CDPparams = New Dictionary
    CDPparams.Add "id", ResultCDP("id")
    Set ResultCDP = controlExtensions.invokeMethod("Extensions.uninstall", CDPparams, True, False)

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

'***************************************************************************************************
'* 機能　　：JavaScript関数、`alert`処理に関するDemoです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：非同期実行、イベントキャプチャした内容をもとにコマンド実行といったことをデモンストレーションします
'* 注意事項：このライブラリのメソッドは、同期前提で組まれてる都合上、低レベル操作で記述します
'***************************************************************************************************
Sub TestAlert()
    '設定シートに基づくブラウザ立ち上げ。`selenium`の独自テストページに遷移します
    Dim Demo_alerts As CDPBrowser: Set Demo_alerts = 設定シートからのCDP起動("https://www.selenium.dev/selenium/web/alerts.html")


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


            ' --- 6. 非同期でコマンド実行(Jsのクリック処理) ---
            'この瞬間、JavaScriptの`alert`関数が発動されます
            Dim AsyncID As Long
            AsyncID = .jsEval("function() { this.click(); }", CStr(resCDP("object")("objectId")), RunAsyncCDP:=True)
    
    
            ' --- 7. イベントキャプチャを有効化 ---
            Set .BrowserEvents = New Dictionary
    
    
            ' --- 8. 特定のイベント名が出るまでループ ---
            Const SearchEventName As String = "Page.javascriptDialogOpening"
            Do
                '非同期イベントを取り出す
                .TakeEvents
    
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
        .quit
    End With
End Sub

'***************************************************************************************************
'* 機能　　：WebView2を使わず、ブラウザそのものを、ExcelUserFromに埋め込み、疑似WebView2を表現します
'---------------------------------------------------------------------------------------------------
'* 詳細説明：WebView2らしさを追及するべく、キオスクモードで立ち上げ、URL遷移のみのユーザーフォームを起動します
'* 注意事項：Edgeへの入力フォーカスが正しく認識できないため現状は、特定領域でのマウスフォーカスで妥協してます
'***************************************************************************************************
Sub ExcelのユーザーフォームにEdgeを埋め込む()
    '1. CDPでEdgeを起動
    Dim 実質WebView2 As CDPBrowser: Set 実質WebView2 = 設定シートからのCDP起動(KioskMode:=fullscreen)
    実質WebView2.navigate "https://github.com/Eschamali/StarterWebScrapingKit"      'このツールのリポジトリURLとして、遷移します

    '2. フォームをロード（まだ表示はしない）
    Load EdgeInExcelForm

    '3. 誘拐（ドッキング）処理を実行させる！
    実質WebView2.sleep  'ちょこっとクールタイム
    If Not (EdgeInExcelForm.AttachEdge(実質WebView2)) Then MsgBox "Edgeのハンドル情報の取得に失敗しました", vbCritical: Exit Sub

    '4. フォームを表示
    EdgeInExcelForm.show

    '5. ブラウザを正常に閉じる
    実質WebView2.quit
End Sub

'***************************************************************************************************
'* 機能　　：リニューアルした`.jsEval`の動作確認用です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：ブラウザ内で直接実行するJavaScriptは様々な返り値が届きます。それに対応できるかの動作確認です
'            ぜひ、ログレベル：DEBUG　にして実行してみてね。
'***************************************************************************************************
Sub jsEval動作確認()
    ' 設定シートに基づくブラウザ立ち上げ
    Dim jsEvalTest As CDPBrowser
    Set jsEvalTest = 設定シートからのCDP起動("https://news.google.com/")

    ' JSON変換用オブジェクト（ログ出力用）
    Dim JsonConv As New WebJsonConverter
    Dim res As Variant

    With jsEvalTest
        Debug.Print "============================================="
        Debug.Print "  jsEval 限界突破 ＆ 新機能テスト 開始"
        Debug.Print "============================================="
        
        Debug.Print vbCrLf & "--- [1] 基本的な型（Primitives） ---"
        res = .jsEval("'Hello VBA!'")            ' String
        Debug.Print res
        
        res = .jsEval("123.45")                  ' Number (Double)
        Debug.Print res
        
        res = .jsEval("true")                    ' Boolean
        Debug.Print res
        
        Debug.Print vbCrLf & "--- [2] JS特有の「無（む）」 ---"
        res = .jsEval("null")                    ' Null (VBAでは Null)
        Debug.Print res
        
        res = .jsEval("undefined")               ' Undefined (VBAでは Empty)
        Debug.Print res
        
        Debug.Print vbCrLf & "--- [3] コレクション（returnByValue = True） ---"
        ' 配列は Dictionary (0,1,2...キー) になる
        Set res = .jsEval("[10, 20, 30]", returnByValue:=True)
        If IsObject(res) Then Debug.Print " Array Count: " & res.Count ' 3
        
        ' オブジェクトは Dictionary (キー名) になる
        Set res = .jsEval("({ name: 'Taro', age: 20 })", returnByValue:=True)
        If IsObject(res) Then Debug.Print " Name: " & res("name")      ' Taro
        
        Debug.Print vbCrLf & "--- [4] DOM要素の参照取得（returnByValue = False） ---"
        ' 値ではなく objectId (住所) を取得する
        Dim bodyId As String
        res = .jsEval("document.querySelector('body')", returnByValue:=False)
        bodyId = CStr(res)
        Debug.Print " Body objectId: " & bodyId
        
        Debug.Print vbCrLf & "--- [5] 新機能！ objectId 指定の callFunctionOn テスト ---"
        ' 取得した bodyId を指定して、その中身のテキストを取得する
        ' ※callFunctionOn の仕様上、function() { ... } 形式で記述する必要があります
        res = .jsEval("function() { return this.tagName;}", objectId:=bodyId, returnByValue:=True)
        Debug.Print " TagName of objectId: " & res ' BODY と出れば大成功！

        Debug.Print vbCrLf & "--- [6] エラー制御（StopError = False） ---"
        ' わざとエラーを発生させ、マクロが止まらずに CVErr が返るかチェック
        res = .jsEval("this_is_error()", StopException:=False)
        If IsError(res) Then
            Debug.Print " 【成功】エラーを検知し、マクロは止まらずに CVErr を返しました！"
        End If
        
        Debug.Print vbCrLf & "--- [7] 新機能！ IFERROR 代替値のテスト ---"
        ' エラー時に、指定した代替値（IFERROR）が返るかチェック
        res = .jsEval("document.querySelector('#unknown_element').innerText", StopException:=False, IFEXCEPTION:="要素が見つかりません")
        Debug.Print " IFERRORの結果: " & res ' 「要素が見つかりません」と出れば大成功！
        
        Debug.Print vbCrLf & "--- [8] 非同期処理（awaitPromise = True） ---"
        ' 1秒待ってから値を返すPromise。VBA側でちゃんと待機できるか？
        res = .jsEval("new Promise(r => setTimeout(() => r('非同期待機、大成功！'), 1000))", awaitPromise:=True)
        Debug.Print " Promise Result: " & res
        
        Debug.Print vbCrLf & "--- [9] 非同期CDP（RunAsyncCDP = True） ---"
        ' VBA側は結果を待たずに即座にIDを返して次へ進む
        res = .jsEval("alert('このアラートはVBAを止めません！')", RunAsyncCDP:=True)
        Debug.Print " Async Command ID: " & res

        Debug.Print vbCrLf & "============================================="
        Debug.Print "  全テスト完了！！"
        Debug.Print "============================================="
    End With

    jsEvalTest.quit
End Sub

'***************************************************************************************************
'* 機能　　：`.jsEval`の上級引数（objectArguments / contextId / serializationOptions）の動作確認用です
'---------------------------------------------------------------------------------------------------
'* 詳細説明：既存の`jsEval動作確認`で検証していない3つの引数を体系的にテストします。
'
'   ① objectArguments   … Runtime.callFunctionOn の `arguments` に相当
'                          callFunctionOn 方式（objectId指定時）で、関数に引数を渡す際に使います。
'                          Collection に Dictionary（{value: xxx} or {objectId: xxx}）を積んで渡します。
'
'   ② contextId         … Runtime.evaluate の実行コンテキストIDを指定します。
'                          主に iFrame 内で JavaScript を実行したいときに使います。
'                          objectId が指定されている場合は無視されます。
'
'   ③ serializationOptions … returnByValue を上書きして、戻り値のシリアライズ方法を細かく制御します。
'                          serialization に "deep" / "json" / "idOnly" を指定できます。
'                          "deep" の場合、maxDepth や additionalParameters（DOM専用オプション等）も使えます。
'
'* 注意事項：・ぜひ、ログレベル：DEBUG　にして実行してみてね。
'            ・contextId の検証には iFrame を含むページが必要なため、w3schools のデモページを使います
'***************************************************************************************************
Sub jsEval高度な引数検証()
    Dim JsonConv As New WebJsonConverter

    Debug.Print "============================================="
    Debug.Print "  jsEval 高度な引数検証テスト 開始"
    Debug.Print "============================================="


    '==========================================================================
    ' ブロック① objectArguments のテスト
    '==========================================================================
    ' objectId 指定（callFunctionOn）方式でのみ有効な引数です。
    ' JavaScript 関数に渡す引数を Collection に積んで渡します。
    ' 各要素は Dictionary 形式 → {value: プリミティブ値} または {objectId: 文字列}
    '--------------------------------------------------------------------------
    Debug.Print vbCrLf & "============================================="
    Debug.Print "  ブロック① objectArguments の検証"
    Debug.Print "============================================="

    Dim browserObjArgs As CDPBrowser
    Set browserObjArgs = 設定シートからのCDP起動("https://news.google.com/")

    Dim bodyObjId As String
    Dim res As Variant

    With browserObjArgs
        ' まず body 要素の objectId を取得しておく（callFunctionOn に渡す対象オブジェクト）
        res = .jsEval("document.querySelector('body')", returnByValue:=False)
        bodyObjId = CStr(res)
        Debug.Print "  body の objectId: " & bodyObjId

        ' --- [① -1] value 型のプリミティブ引数を渡す ---
        ' テスト内容: 関数に数値を2つ渡して足し算した結果を返す
        ' 期待結果  : 30 (10 + 20)
        Debug.Print vbCrLf & "--- [①-1] プリミティブ引数（数値） 渡しテスト ---"
        Dim args1 As New Collection
        Dim arg1a As New Scripting.Dictionary, arg1b As New Scripting.Dictionary
        arg1a.Add "value", 10
        arg1b.Add "value", 20
        args1.Add arg1a
        args1.Add arg1b

        ' callFunctionOn 方式では、関数内 arguments[0], arguments[1] ... で受け取れます
        res = .jsEval("function(a, b) { return a + b; }", objectId:=bodyObjId, _
                      objectArguments:=args1, returnByValue:=True)
        Debug.Print "  引数: 10 + 20 = " & res & IIf(res = 30, "  ←大成功！", "  ←想定外")


        ' --- [①-2] String 型の引数を渡す ---
        ' テスト内容: 渡した文字列を大文字に変換して返す
        ' 期待結果  : "HELLO FROM VBA"
        Debug.Print vbCrLf & "--- [①-2] プリミティブ引数（文字列） 渡しテスト ---"
        Dim args2 As New Collection
        Dim arg2a As New Scripting.Dictionary
        arg2a.Add "value", "Hello from vba"
        args2.Add arg2a

        res = .jsEval("function(str) { return str.toUpperCase(); }", objectId:=bodyObjId, _
                      objectArguments:=args2, returnByValue:=True)
        Debug.Print "  結果: " & res & IIf(res = "HELLO FROM VBA", "  ←大成功！", "  ←想定外")


        ' --- [①-3] objectId 型の引数を渡す ---
        ' テスト内容: 既取得の body objectId を改めて引数として渡し、タグ名を返す
        ' 期待結果  : "BODY"
        ' ※ objectId を引数に渡す場合は {objectId: "..."} 形式で Collection に追加します
        Debug.Print vbCrLf & "--- [①-3] objectId 型引数 渡しテスト ---"
        Dim args3 As New Collection
        Dim arg3a As New Scripting.Dictionary
        arg3a.Add "objectId", bodyObjId
        args3.Add arg3a

        ' this（bodyObjId を指定したオブジェクト）context 上で実行するため、
        ' 引数で渡した要素のタグ名を参照できるか確認します
        res = .jsEval("function(el) { return el.tagName; }", objectId:=bodyObjId, _
                      objectArguments:=args3, returnByValue:=True)
        Debug.Print "  結果: " & res & IIf(res = "BODY", "  ←大成功！", "  ←想定外")

        .quit
    End With


    '==========================================================================
    ' ブロック② contextId のテスト
    '==========================================================================
    ' iFrame 内の JavaScript 実行コンテキストIDを指定して実行します。
    ' まず Runtime.enable を有効にし executionContextCreated イベントで contextId を取得します。
    ' 取得した contextId を指定することで、iFrame 内のDOMを直接操作できます。
    '--------------------------------------------------------------------------
    Debug.Print vbCrLf & "============================================="
    Debug.Print "  ブロック② contextId の検証"
    Debug.Print "============================================="

    ' iFrame を含む公開デモページを使用します（W3Schools）
    Dim browserCtxId As CDPBrowser
    Set browserCtxId = 設定シートからのCDP起動

    Dim ResultCDP As Dictionary
    Dim parsCDP As New Scripting.Dictionary

    With browserCtxId
        ' まず外側から通常 evaluate し、メインコンテキストで動いていることを確認
        Debug.Print vbCrLf & "--- [②-1] メインコンテキストでの実行確認 ---"
        res = .jsEval("window.location.href")
        Debug.Print "  現在URL（メインコンテキスト）: " & res

        ' Runtime を有効化して executionContextCreated イベントを拾えるようにする
        Debug.Print vbCrLf & "--- [②-2] contextId の列挙 ---"

        ' イベントキャプチャを有効化
        Set .BrowserEvents = New Dictionary
       .navigate "https://www.w3schools.com/html/tryit.asp?filename=tryhtml_iframe_height_width"

        ' 少し待ってからイベントを取り出す（ページロード後のコンテキスト情報をキャプチャ）
        .sleep 1
        .TakeEvents

        ' executionContextCreated イベントから contexId を収集する
        Dim iframeContextId As Long
        iframeContextId = 0
        If .BrowserEvents("EventMethods").Exists("Runtime.executionContextCreated") Then
            Dim ctx As Variant
            For Each ctx In .BrowserEvents("EventMethods")("Runtime.executionContextCreated")
                Dim ctxDesc As Dictionary
                Set ctxDesc = ctx("params")("context")
                Debug.Print "  contextId: " & ctxDesc("id") & " | origin: " & ctxDesc("origin") & " | name: " & ctxDesc("name")

                ' iFrame のコンテキストを見分ける（name が空でなく、id > 1 のものを使う）
                If ctxDesc("id") > 1 And iframeContextId = 0 Then
                    iframeContextId = CLng(ctxDesc("id"))
                    Debug.Print "    ↑ このcontextIdをiFrameテスト用に使用します"
                End If
            Next
        Else
            Debug.Print "  ※ executionContextCreated イベントが取れませんでした。Runtime.enable の前にすでに存在していた可能性があります"
        End If

        ' Runtime.enable で改めてコンテキスト一覧を取得する別手段も試す
        ' → Runtime.getFrameTree からの contextId 照合 等が必要な場合もある（参考用コメント）

        ' イベントキャプチャを無効化
        Set .BrowserEvents = Nothing
        .invokeMethod "Runtime.disable"


        ' --- [②-3] 取得した contextId で JavaScript を実行 ---
        If iframeContextId > 0 Then
            Debug.Print vbCrLf & "--- [②-3] iFrame contextId(" & iframeContextId & ") 内での JS 実行テスト ---"
            ' そのコンテキストで window.location.href を取得することで、iFrame内のURLが返るはず
            res = .jsEval("window.location.href", contextId:=iframeContextId, StopException:=False)
            Debug.Print "  contextId=" & iframeContextId & " での URL: " & res
            If VarType(res) = vbString And InStr(res, "http") > 0 Then
                Debug.Print "  ←iFrame コンテキスト内での実行に成功！"
            Else
                Debug.Print "  ←結果を確認してください（iFrame URLが取れていない可能性があります）"
            End If
        Else
            Debug.Print "  ※ iFrame の contextId が取得できなかったため、[②-3] はスキップします"
            Debug.Print "    ヒント: ページを開く前に Runtime.enable を有効にして、contextCreated をキャプチャする必要があります"
        End If

        .quit
    End With


    '==========================================================================
    ' ブロック③ serializationOptions のテスト
    '==========================================================================
    ' returnByValue を上書きして、戻り値のシリアライズ方法を細かく制御します。
    ' serialization フィールドに以下を指定できます:
    '   "json"    → returnByValue:=True と同等（JSONシリアライズ）
    '   "deep"    → 深いオブジェクトもシリアライズ。maxDepth で深さ制御
    '   "idOnly"  → objectId のみ返す（returnByValue:=False と同等）
    '--------------------------------------------------------------------------
    Debug.Print vbCrLf & "============================================="
    Debug.Print "  ブロック③ serializationOptions の検証"
    Debug.Print "============================================="

    Dim browserSerial As CDPBrowser
    Set browserSerial = 設定シートからのCDP起動("https://news.google.com/")

    With browserSerial

        ' --- [③-1] serialization: "json" ---
        ' テスト内容: returnByValue:=True と同等の動作をするか確認
        ' 期待結果  : Dictionary として {name: "Taro", age: 20} が返る
        Debug.Print vbCrLf & "--- [③-1] serialization: ""json"" テスト ---"
        Dim opts1 As New Scripting.Dictionary
        opts1.Add "serialization", "json"

        Set res = .jsEval("({ name: 'Taro', age: 20 })", serializationOptions:=opts1)
        If IsObject(res) Then
            Debug.Print "  name: " & res("name") & " / age: " & res("age") & "  ←Dictionary で取得成功！"
        Else
            Debug.Print "  結果: " & res & "  ←Dictionary を期待していましたが、違う型で返りました"
        End If


        ' --- [③-2] serialization: "idOnly" ---
        ' テスト内容: returnByValue:=False と同等。objectId のみ返す
        ' 期待結果  : objectId 文字列（"-xxxx.x.x" 形式）が String で返る
        Debug.Print vbCrLf & "--- [③-2] serialization: ""idOnly"" テスト ---"
        Dim opts2 As New Scripting.Dictionary
        opts2.Add "serialization", "idOnly"

        res = .jsEval("document.querySelector('body')", serializationOptions:=opts2)
        Debug.Print "  結果: " & res
        If VarType(res) = vbString And Len(CStr(res)) > 0 Then
            Debug.Print "  ←objectId 取得成功！（再利用可能なID）"
        Else
            Debug.Print "  ←想定外の型です"
        End If


        ' --- [③-3] serialization: "deep" with maxDepth ---
        ' テスト内容: ネストされたオブジェクトを指定深さまでシリアライズして返す
        ' 期待結果  : Dictionary（またはそれに準ずる）として返る。深さ3まで展開される
        ' ※ "deep" は DeepSerializedValue として返るため、RemoteObject の deepSerializedValue フィールドに入ります
        '    （jsEval の JSresultAnalysis では、通常 value フィールドがない場合 description を返します）
        Debug.Print vbCrLf & "--- [③-3] serialization: ""deep"" with maxDepth テスト ---"
        Dim opts3 As New Scripting.Dictionary
        opts3.Add "serialization", "deep"
        opts3.Add "maxDepth", 3

        Set res = .jsEval("({ level1: { level2: { level3: 'deep value' } } })", serializationOptions:=opts3)
        If IsObject(res) Then
            Debug.Print "  結果(Dictionary形式): " & JsonConv.ConvertToJson(res) & "  ←  成功！"
        Else
            Debug.Print "  結果: " & CStr(res) & "  ← 内容を確認してください（deepSerializedValue に入ってきている可能性あり）"
        End If


        ' --- [③-4] serialization: "deep" + DOM additionalParameters（Chromeでの DOM 追加オプション） ---
        ' テスト内容: body 要素を deep シリアライズ。DOM専用の maxNodeDepth と includeShadowTree を追加
        ' 期待結果  : type:"node" として DeepSerializedValue 形式の情報が返る
        ' ※ additionalParameters は Chrome での DOM シリアライズ用拡張パラメーター（任意）
        Debug.Print vbCrLf & "--- [③-4] serialization: ""deep"" + DOM additionalParameters テスト ---"
        Dim opts4 As New Scripting.Dictionary
        Dim addParams As New Scripting.Dictionary
        addParams.Add "maxNodeDepth", 2
        addParams.Add "includeShadowTree", "none"
        opts4.Add "serialization", "deep"
        opts4.Add "maxDepth", 2
        opts4.Add "additionalParameters", addParams

        Set res = .jsEval("document.querySelector('body')", serializationOptions:=opts4)
        If IsObject(res) Then
            Debug.Print "  結果(body DOM deep): " & JsonConv.ConvertToJson(res) & "  ← 成功！"
        Else
            Debug.Print "  結果: " & CStr(res) & " 内容を確認してください"
        End If

        .quit
    End With

    Debug.Print vbCrLf & "============================================="
    Debug.Print "  全ブロックのテスト完了！！"
    Debug.Print "============================================="
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
    Dim edge As CDPBrowser
    Set edge = 設定シートからのCDP起動
 
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

    Dim chrome As CDPBrowser
 
   'Start and hide
    Set chrome = 設定シートからのCDP起動
    chrome.hide
 
   'Perform automation in the background
    chrome.navigate "https://google.com", isInteractive
    chrome.getElementByQuery("[name='q']").value = "automate edge vba"
    chrome.getElementByQuery("[name='q']").submit
    
   'Click the target result link
    chrome.getElementByXPath("//h3[text()='Automate Chrome / Edge using VBA']").click
    
   'Get the vote count only once the target element appears on screen
   'The onExists method is needed as this element appears after ReadyState = "complete"
    Dim voteCount As Long
    voteCount = chrome.getElementByID("ctl00_RateArticle_VoteCountNoHist").onExist.innerHTML
    
   'Confirm result and display
    Dim userChoice
    userChoice = MsgBox("Automation completed. Current vote counts: " & voteCount & ". Do you want to see the window?", vbYesNo)
    If userChoice = vbYes Then chrome.show Else chrome.quit
    
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

    Dim chrome As CDPBrowser

    'Start and hide
    Set chrome = 設定シートからのCDP起動
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
 
    Dim chrome As CDPBrowser
    Set chrome = 設定シートからのCDP起動
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
 
    Dim chrome As CDPBrowser
    Set chrome = 設定シートからのCDP起動
    chrome.show
 
   'Create and assign tabs
    Dim tab1 As New CDPBrowser                   'The keyword "New" is a must
    Dim tab2 As New CDPBrowser
    Dim tab3 As New CDPBrowser
    Set tab1 = chrome                            'The first tab is open by default after .start
    Set tab2 = chrome.newTab(newWindow:=True)    'newWindow: open tab as a new window instead of a tab
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
    Dim chrome As CDPBrowser
    Set chrome = 設定シートからのCDP起動
    'chrome.start addArgs:="--disable-popup-blocking"    'The disable-popup-blocking argument is needed to allow opening link in a new tab
    chrome.show asMaximized
    
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
    Dim targetTab As New CDPBrowser
    Set targetTab = chrome.getTab
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
    
    Dim chrome As New CDPBrowser
    Set chrome = 設定シートからのCDP起動(demoUrl)
    
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
    
    Dim chrome As CDPBrowser
    Set chrome = 設定シートからのCDP起動   'not App Mode as sometimes Chrome App Mode does not allow file downloading
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
    
    Dim chrome As CDPBrowser
    Set chrome = 設定シートからのCDP起動
    chrome.navigate demoUrl
        
   'Get the target fields
    Dim ip As CDPElement
    Dim sb As CDPElement
    Set ip = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='text']")
    Set sb = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='submit']")
        
   'This traditional input method will fail as this is a React field
    chrome.jsEval ip.varName & ".value = 'TEST1'"
    chrome.jsEval ip.varName & ".dispatchEvent(new Event('input', { bubbles: true, simulated: true }))"
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

    Dim chrome As CDPBrowser
    Set chrome = 設定シートからのCDP起動
    chrome.newTab "http://google.com", setMain:=True   'the chrome object will now directly refer to the Google tab
    chrome.getTab("about:blank").closeTab       'prior 2.7, the next line will throw an error due to no main-switching mechanism
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
    
    Dim e1 As CDPBrowser
    Set e1 = 設定シートからのCDP起動
    e1.navigate "https://yahoo.com"
    
    Debug.Print Format(Now, "hh:mm:ss") & " execBot1 completed."

End Function

Function execBot2()
'----------------------------------------------------------------------------------------
' Refer to the demoMultiProfileOperation
'----------------------------------------------------------------------------------------

    Debug.Print Format(Now, "hh:mm:ss") & " execBot2 started."

    Dim e2 As CDPBrowser
    Set e2 = 設定シートからのCDP起動(, "CDP2")
    e2.navigate "https://finance.yahoo.com"
    
    Debug.Print Format(Now, "hh:mm:ss") & " execBot2 completed."

End Function

Sub demoReattachmentPart1()
'----------------------------------------------------------------------------------------
' From v3.1, .reattach is necessary to perform reattachment to the current CDP instances
' as each instance is now identified with a unique user profile for multi-instances
' operation. The below procedure starts a new CDP session under profile CDP2. After
' running demoReattachmentPart1, you can run demoReattachmentPart2 to see the correct
' way of applying .reattach to the CDP2 session.
'----------------------------------------------------------------------------------------

    Dim c As CDPBrowser
    Set c = 設定シートからのCDP起動
    c.navigate "https://google.com"

End Sub

Sub demoReattachmentPart2()
'----------------------------------------------------------------------------------------
' See notes in demoReattachmentPart1
'----------------------------------------------------------------------------------------

    Dim c As New CDPBrowser

    '設定セルから、ユーザ名を取得
    With ShSetting01_StartBrowser
        Dim UserName As String
        UserName = .Range(.UseRangeName(2, "Demo_CDP.demoReattachmentPart2")).value
    End With

    '1. まずは、既存のTargetIDに接続できるか？
    If c.reattach(UserName) Then
        '接続できたので、別ページに遷移して終了
        c.navigate "https://wikipedia.com"
        Exit Sub
    Else
        '既存のTargetIDが消えちゃったので、次のフェーズへ
        Debug.Print "Failed to reattach. Connecting to the nearest unconnected tab from `Target.getTargets`."
    End If

    '2. 最も近い未接続のタブに接続します
    c.getTab setMain:=True

    '3．再接続できたので、別ページに遷移して終了
    c.navigate "https://wikipedia.com"
End Sub



'***************************************************************************************************
'                               ■■■ ヘルパープロシージャ ■■■
'***************************************************************************************************
'* 機能　　：このExcelが、OneDrive上で実行されてる場合のパス変換処理を行います
'---------------------------------------------------------------------------------------------------
'* 返り値　：ローカルパス
'* 引数　　：Path                   基本は、`thisworkbook.path`を指定
'            UsePrivateOneDrive     社内個人OneDriveの場合は、`False`にしてください
'---------------------------------------------------------------------------------------------------
'* 機能説明：開いてるExcelがOneDriveにあると、`thisworkbook.path`がインターネット上のURLになってしまい、一部操作ができなくなる問題に対処した物となります。
'            純ローカルなら、そのまま返します。
'            個人向けOneDrive と ビジネス向け個人OneDrive に対応してます。先頭の定数で、スイッチングしてください
'
'* 注意事項：SharePointの場合は、自力でコードを書く必要があります
'***************************************************************************************************
Function OneDrivePathToLocalPath(Path As String, Optional UsePrivateOneDrive As Boolean = True) As String
    'http始まりじゃないなら、そのまま返して終了
    If Left(Path, 4) <> "http" Then OneDrivePathToLocalPath = Path: Exit Function

    '個人OneDriveモードなら識別番号分、ローカルパスに置き換えて結合
    If UsePrivateOneDrive Then
        OneDrivePathToLocalPath = Environ("OneDrive") & Mid(Path, 41)
    
    '個人BusinessOneDriveモードなら"Documents"以降のパスを抜き出して、ローカルパスと結合
    Else
        OneDrivePathToLocalPath = Environ("OneDriveCommercial") & Evaluate("TEXTAFTER(""" & Path & """,""/Documents"")")
    End If
End Function
