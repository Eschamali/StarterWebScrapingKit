Attribute VB_Name = "Demo_BiDi"
'==============================================================================================================
'               Automating Chromium-Based Browsers with WebDriverBiDi API and VBA
'--------------------------------------------------------------------------------------------------------------
'
'==============================================================================================================
Option Explicit

    
'--------------------------------------------------------------------------------------------------------------
' Module      : BiDiDemo
' Description : BiDiCore.cls を用いて、ChromiumブラウザでのWebDriver BiDi通信を確認するデモプログラム。
'               BiDiPoC.basの内容を、新設したBiDiCore.clsを利用してリファクタリングしたものです。
'--------------------------------------------------------------------------------------------------------------

Public Sub TestBiDiCoreDemo()
    Dim jsConverter As New WebJsonConverter

    
    ' 2. BiDiCore の初期化
    Dim bidi As New BiDiCore
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
    
    'BiDiCoreの初期化とブラウザ立ち上げ
    Dim Demo_NetworkEvent As New BiDiCore
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
