Attribute VB_Name = "Demo_BiDi"
Option Explicit

Private Declare PtrSafe Sub sleep Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)
    
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
