Attribute VB_Name = "Demo_BiDi"
Option Explicit

'--------------------------------------------------------------------------------------------------------------
' Module      : BiDiDemo
' Description : BiDiCore.cls を用いて、ChromiumブラウザでのWebDriver BiDi通信を確認するデモプログラム。
'               BiDiPoC.basの内容を、新設したBiDiCore.clsを利用してリファクタリングしたものです。
'--------------------------------------------------------------------------------------------------------------

Public Sub TestBiDiCoreDemo()
    Dim jsConverter As New WebJsonConverter


    ' 1. CDPBrowserの起動
    Dim targetBrowser As CDPBrowser: Set targetBrowser = 設定シートからの起動
    Set targetBrowser.BrowserEvents = New Dictionary
    
    ' 初期URLとして適当な空ページへ
    targetBrowser.navigate "about:blank#MAPPER_TARGET", isComplete

    ' 現在のターゲットIDの取得 (CDPBrowserのプロパティまたは内部から取得する想定)
    ' ※CDPBrowser.getTab() などで取得できる内部ターゲットIDが必要です
    Dim current_targetID As String
    ' テスト用に既存の deserializeFromTable 機構を呼ぶ例（実際のプロジェクトの作りに合わせて調整可）
    current_targetID = ShSetting01_StartBrowser.deserializeFromTable("Automation Data", 2, "TestBiDi")("targetID")
    
    ' 2. BiDiCore の初期化
    Dim bidi As New BiDiCore
    bidi.Init targetBrowser, current_targetID
    
    ' あとで遷移を確認するための新しいタブをCDP経由で開く
    targetBrowser.newTab
    
    Dim result As Dictionary
    Dim params As Dictionary
    Dim tmp As Variant
    
    ' 3. BiDiコマンド [session.new] の送信
    Debug.Print "--- Sending session.new ---"
    Set params = New Dictionary
    params.Add "capabilities", New Dictionary
    Set result = bidi.invokeMethod("session.new", params)
    
    Debug.Print "session.new result:"
    Debug.Print jsConverter.ConvertToJson(result)
    
    ' 4. BiDiコマンド [browsingContext.getTree] の送信
    Debug.Print "--- Sending browsingContext.getTree ---"
    Set result = bidi.invokeMethod("browsingContext.getTree")
    
    Debug.Print "browsingContext.getTree result:"
    Debug.Print jsConverter.ConvertToJson(result)
    
    ' 取得したコンテキストの1つをターゲットとして抽出する例
    Dim targetContext As String
    If result.Exists("result") Then
        If result("result").Exists("contexts") Then
            If result("result")("contexts").Count > 0 Then
                ' 例として配列の2番目（インデックス2、VBAは1列目からかもしれないので適切に抽出）
                ' ここでは雑に最初のcontextを取得します。
                targetContext = result("result")("contexts")(1)("context")
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
    Debug.Print "Subscribe result:"
    Debug.Print jsConverter.ConvertToJson(result)
    
    Debug.Print ">>> Waiting for console.log in the browser... (type console.log('Hello BiDi!'))"
    
    ' 7. 非同期イベントの待機
    Dim evName As String
    evName = "log.entryAdded"
    ' 無限ループでイベントを捕まえるデモ
    Do
        bidi.TakeEvents
        
        If bidi.Events.Exists(evName) Then
            Debug.Print "--- " & evName & " Event triggerd! ---"
            
            Dim loggedEvents As Collection
            Set loggedEvents = bidi.Events(evName)
            
            For Each tmp In loggedEvents
                Debug.Print jsConverter.ConvertToJson(tmp)
            Next
            
            ' 取得後はキューから消す
            bidi.Events.Remove evName
            Exit Do
        End If
        
        targetBrowser.sleep
        DoEvents
    Loop While True
    
    Debug.Print "--- BiDiDemo Finished ---"
    
    ' クリーンアップ
    targetBrowser.quit
End Sub
