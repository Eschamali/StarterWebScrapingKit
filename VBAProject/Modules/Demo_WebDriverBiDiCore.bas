Attribute VB_Name = "Demo_WebDriverBiDiCore"
'==============================================================================================================
'               Automating Chromium-Based Browsers with WebDriverBiDi API and VBA
'--------------------------------------------------------------------------------------------------------------
'
'==============================================================================================================
Option Explicit



'--------------------------------------------------------------------------------------------------------------
' Module      : Demo_WebDriverBiDiCore
' Description : WebDriverBiDiCore.cls を用いて、ChromiumブラウザでのWebDriver BiDi通信を確認するデモプログラム。
'               BiDiPoC.basの内容を、新設したWebDriverBiDiCore.clsを利用してリファクタリングしたものです。
'--------------------------------------------------------------------------------------------------------------

Public Sub TestWebDriverBiDiCoreDemo()
    Dim jsConverter As New WebJsonConverter

    
    ' 2. WebDriverBiDiCore の初期化
    Dim bidi As New WebDriverBiDiCore
    bidi.start
    
    
    Dim result As Dictionary
    Dim params As Dictionary
    Dim tmp As Variant
    
    
    ' 4. BiDiコマンド [browsingContext.getTree] の送信
    Debug.Print "--- Sending browsingContext.getTree ---"
    Set result = bidi.invokeMethod("browsingContext.getTree")
    
    Debug.Print "browsingContext.getTree result:"
    Debug.Print jsConverter.ConvertToJson(result)
    
    ' 取得したコンテキストの1つをターゲットとして抽出する例
    Dim targetContext As String
    If Not (result Is Nothing) Then
        If result.Exists("contexts") Then
            If result("contexts").Count > 0 Then
                ' 例として配列の2番目（インデックス2、VBAは1列目からかもしれないので適切に抽出）
                ' ここでは雑に最初のcontextを取得します。
                targetContext = result("contexts")(1)("context")
            End If
        End If
    End If
    
    If targetContext <> "" Then
        ' 5. コンテキストのナビゲート [browsingContext.navigate]
        Debug.Print "--- Navigating context: " & targetContext & " ---"
        Set params = New Dictionary
        params.Add "context", targetContext
        params.Add "url", "https://www.google.com"
        params.Add "wait", "complete"
        
        Set result = bidi.invokeMethod("browsingContext.navigate", params)
        Debug.Print "Navigation result:"
        Debug.Print jsConverter.ConvertToJson(result)
    Else
        Debug.Print "Could not extract a target context from getTree."
    End If
    
    ' 6. イベント付きのサブスクライブ [session.subscribe]
    Debug.Print "--- Subscribing to log.entryAdded ---"
    Set params = New Dictionary
    Dim eventsArray As New Collection
    eventsArray.Add "log.entryAdded"
    params.Add "events", eventsArray
    
    Set result = bidi.invokeMethod("session.subscribe", params)
    Set bidi.BiDiEvents = New Dictionary
    Debug.Print "Subscribe result:"
    Debug.Print jsConverter.ConvertToJson(result)
    
    Debug.Print ">>> Waiting for console.log in the browser... (type console.log('Hello BiDi!'))"
    
    ' 7. 非同期イベントの待機
    Dim evName As String
    evName = "log.entryAdded"
    ' 無限ループでイベントを捕まえるデモ
    Do
        bidi.TakeEvents
        
        If bidi.BiDiEvents("EventMethods").Exists(evName) Then
            Debug.Print "--- " & evName & " Event triggerd! ---"
            
            Dim loggedEvents As Collection
            Set loggedEvents = bidi.BiDiEvents("EventMethods")(evName)
            
            For Each tmp In loggedEvents
                Debug.Print jsConverter.ConvertToJson(tmp)
            Next
            
            ' 取得後はキューから消す
            Set bidi.BiDiEvents = New Dictionary
            Exit Do
        End If
        
        bidi.sleep
        DoEvents
    Loop While True
    
    Debug.Print "--- BiDiDemo Finished ---"

    ' クリーンアップ
    bidi.quit
End Sub

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
    Dim Demo_NetworkEvent As New WebDriverBiDiCore
    Demo_NetworkEvent.start
    
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
    Demo_alerts.start "https://www.selenium.dev/selenium/web/alerts.html", , caps

    '結果とBiDiパラメーター変数を用意
    Dim paramsBiDi As Dictionary, resultBiDi As Dictionary

    '現在のコンテキストIDを取得する
    Set resultBiDi = Demo_alerts.invokeMethod("browsingContext.getTree")
    Dim targetContext As String
    If Not (resultBiDi Is Nothing) Then
        targetContext = resultBiDi("contexts")(1)("context")    '一旦は、先頭ブラウザで　※本来はURLcheckとかがいると思うが、低レベル制御の都合上、妥協
    End If

'    'ページ遷移の場合のコード
'    Set paramsBiDi = New Dictionary
'    paramsBiDi.Add "context", targetContext
'    paramsBiDi.Add "url", "https://www.selenium.dev/selenium/web/alerts.html"
'    paramsBiDi.Add "wait", "complete"
'    Demo_alerts.invokeMethod "browsingContext.navigate", paramsBiDi

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
            Set resBiDiAsync = .ResultBiDiForAsync(AsyncID)
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
