Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                   WebSocket のデモンストレーションです（非同期モードのみ）
'
'   SafeTimer 方式のコールバック統一により、AddressOf / isDataReady は不要になりました。
'   コールバックは WebSocketCommunicator.Http_OnCallback で自動的に受け取られます。
'***************************************************************************************************
Option Explicit

Private g_WebsocketObj As WebSocketCommunicator
Private g_CdpRequestId As Long

Private Function CdpNextRequestId() As Long
    g_CdpRequestId = g_CdpRequestId + 1
    CdpNextRequestId = g_CdpRequestId
End Function


'***************************************************************************************************
'                       ■■■ 非同期処理(echo.websocket.org 編) ■■■
'***************************************************************************************************
'* 機能　　：wss://echo.websocket.org に新規非同期接続し、送受信テストをします
'---------------------------------------------------------------------------------------------------
'* 注意事項：接続後、受信予約を行うとコールバック経由でデータが返ります
'***************************************************************************************************
Sub WebSocketDemoASync1_1_初期化()
    'オブジェクトを作成（SafeTimer 方式のコールバックは Init 内で自動登録される）
    Set g_WebsocketObj = New WebSocketCommunicator

    '接続先を設定します（AddressOf 不要）
    Dim ResultHandleCode As LongPtr: ResultHandleCode = g_WebsocketObj.Init("echo.websocket.org", "")

    '成功判定
    If ResultHandleCode Then
        Dim g_ReConnectionHandle As LongPtr
        g_ReConnectionHandle = ResultHandleCode
        Debug.Print "Websocket connect is success. AsyncMode."
        Debug.Print "ハンドルコード：" & ResultHandleCode

        '1件分の送信をしてみる
        Dim ResultCode As Long: ResultCode = g_WebsocketObj.SendMessage("うみねこ！みゃ～お！" & WorksheetFunction.Unichar(129418))

        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then
            Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If

        '受信予約を行う（コールバック経由で Http_OnCallback が発火し m_isReceiving がセットされる）
        Debug.Print g_WebsocketObj.GetAsyncMessage(, ResultCode)

        If ResultCode Then
            Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print "受信結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub

Sub WebSocketDemoASync1_2_受信リクエスト()
    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync1_1_初期化 を実行してください。"
        Exit Sub
    End If

    '受信メッセージを受け取る
    Dim ResultCode As Long
    g_WebsocketObj.WsLogInfo "受信内容：" & g_WebsocketObj.GetAsyncMessage(, ResultCode), "Demo"

    Dim ErrorMes As New WinApiError
    If ResultCode Then
        g_WebsocketObj.WsLogInfo "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"), "Demo"
    Else
        g_WebsocketObj.WsLogInfo "受信結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp"), "Demo"
    End If
End Sub

Sub WebSocketDemoASync1_3_ハンドルから送信()
    'カウント用
    Static Count As Long
    Count = Count + 1

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync1_1_初期化 を実行してください。"
        Exit Sub
    End If

    '1件分の送信をしてみる
    Dim ResultCode As Long: ResultCode = g_WebsocketObj.SendMessage("うみねこ！みゃ～お！" & Count & WorksheetFunction.Unichar(129418))
'        Dim ResultCode As Long: ResultCode = g_WebsocketObj.SendMessage("{""id"":" & 1 & "," & _
'                  """method"":""Network.getAllCookies""," & _
'                  """params"":{}}")

    Dim ErrorMes As New WinApiError
    If ResultCode Then
        Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "送信結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If
End Sub



'***************************************************************************************************
'                       ■■■ 非同期処理(chrome devtools Protocol 編) ■■■
'***************************************************************************************************
'* 機能　　：CDP（Chrome DevTools Protocol）を WebSocket 経由で叩くデモ
'---------------------------------------------------------------------------------------------------
'* 詳細説明：
'*   2_1 初期化 / 2_2 Runtime.evaluate / 2_3 受信 / 2_4 後始末
'*   2_5 Network.getAllCookies（長文レスポンス）→ 続けて 2_3 で受信
'*   2_6 Page.navigate（ブラウザが遷移する）
'*   2_7 Page.captureScreenshot（スクリーンショット。忘れがちなコマンドはこれ）
'*   2_8 シナリオ：遷移 → 少し待機 → スクショ（中身は 2_3 で受信）
'***************************************************************************************************
Sub WebSocketDemoASync2_1_初期化()
    Const CDP_HOST As String = "127.0.0.1"
    Const CDP_PORT As Long = 9222
    Const CDP_TARGET_PATH As String = "devtools/page/6125BBE810D35B2F7372345B68E7E653"
    Dim ResultHandleCode As LongPtr
    Dim ResultCode As Long
    Dim ErrorMes As New WinApiError

    ' Chrome は下記のように起動しておく必要があります:
    ' chrome.exe --remote-debugging-port=9222
    ' その後、http://127.0.0.1:9222/json を開いて webSocketDebuggerUrl の page id を確認し、
    ' CDP_TARGET_PATH の REPLACE_WITH_TARGET_ID を置き換えてください。

    Set g_WebsocketObj = New WebSocketCommunicator
    ResultHandleCode = g_WebsocketObj.Init(CDP_HOST, CDP_TARGET_PATH, CDP_PORT, False)

    If ResultHandleCode = 0 Then
        Debug.Print "CDP接続失敗。target id の設定を確認してください。"
        Exit Sub
    End If

    Debug.Print "CDP WebSocket connect is success. AsyncMode."
    Debug.Print "ハンドルコード：" & ResultHandleCode
    g_CdpRequestId = 0

    ' 接続直後に受信予約だけ張っておく
    Debug.Print "受信予約：" & g_WebsocketObj.GetAsyncMessage(, ResultCode)
    If ResultCode Then
        Debug.Print "受信予約エラー。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信予約結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If
End Sub

Sub WebSocketDemoASync2_2_CDP送信_RuntimeEvaluate()
    Dim ResultCode As Long
    Dim Payload As String
    Dim rid As Long
    Dim ErrorMes As New WinApiError

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync2_1_初期化 を実行してください。"
        Exit Sub
    End If

    rid = CdpNextRequestId()
    Payload = "{""id"":" & CStr(rid) & ",""method"":""Runtime.evaluate"",""params"":{""expression"":""document.title"",""returnByValue"":true}}"

    ResultCode = g_WebsocketObj.SendMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & rid & ", Runtime.evaluate)"
    End If
End Sub

Sub WebSocketDemoASync2_3_CDP受信()
    Dim ResultCode As Long
    Dim ResponseText As String
    Dim ErrorMes As New WinApiError

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync2_1_初期化 を実行してください。"
        Exit Sub
    End If

    ResponseText = g_WebsocketObj.GetAsyncMessage(, ResultCode)
    Call CdpDebugPrintReceived(ResponseText, ResultCode, ErrorMes)
End Sub

Private Sub CdpDebugPrintReceived(ByVal ResponseText As String, ByVal ResultCode As Long, ByVal ErrorMes As WinApiError)
    Dim previewLen As Long
    Const MAX_PREVIEW As Long = 400

    If ResultCode Then
        Debug.Print "CDP受信エラー。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Exit Sub
    End If

    Debug.Print "CDP受信結果(WinHttp)：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
    Debug.Print "CDP受信文字数：" & Len(ResponseText)
    previewLen = Len(ResponseText)
    If previewLen > MAX_PREVIEW Then
        Debug.Print "CDP受信先頭 " & MAX_PREVIEW & " 文字：" & Left$(ResponseText, MAX_PREVIEW) & " ... (省略)"
    Else
        Debug.Print "CDP受信全文：" & ResponseText
    End If
End Sub

Sub WebSocketDemoASync2_4_後始末()
    If g_WebsocketObj Is Nothing Then Exit Sub

    If g_WebsocketObj.CloseWebSocket(True) Then
        Debug.Print "CDP WebSocket close success."
    Else
        Debug.Print "CDP WebSocket close failed."
    End If

    Set g_WebsocketObj = Nothing
    g_CdpRequestId = 0
End Sub

'* Network.getAllCookies → クッキー一覧の長い JSON（長文レスポンスの負荷テスト向け）
Sub WebSocketDemoASync2_5_CDP_Network_GetAllCookies()
    Dim ResultCode As Long
    Dim Payload As String
    Dim rid As Long
    Dim ErrorMes As New WinApiError

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync2_1_初期化 を実行してください。"
        Exit Sub
    End If

    rid = CdpNextRequestId()
    Payload = "{""id"":" & CStr(rid) & ",""method"":""Network.getAllCookies"",""params"":{}}"
    ResultCode = g_WebsocketObj.SendMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー(Network.getAllCookies)。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & rid & ", Network.getAllCookies) → 続けて 2_3 で受信（長い場合は 2_3 を複数回）"
    End If
End Sub

'* Page.navigate → 実際にタブの URL が変わる
Sub WebSocketDemoASync2_6_CDP_Page_Navigate(Optional ByVal TargetUrl As String)
    Dim ResultCode As Long
    Dim Payload As String
    Dim rid As Long
    Dim esc As String
    Dim ErrorMes As New WinApiError

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync2_1_初期化 を実行してください。"
        Exit Sub
    End If

    If Len(TargetUrl) = 0 Then TargetUrl = "https://www.wikipedia.org/"
    esc = CdpJsonEscape(TargetUrl)

    rid = CdpNextRequestId()
    Payload = "{""id"":" & CStr(rid) & ",""method"":""Page.navigate"",""params"":{""url"":""" & esc & """}}"
    ResultCode = g_WebsocketObj.SendMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー(Page.navigate)。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & rid & ", Page.navigate url=" & TargetUrl & ") → ブラウザの表示を確認し、必要なら 2_3 で受信"
    End If
End Sub

'* Page.captureScreenshot → result.data に base64 PNG（長文になりがち）
Sub WebSocketDemoASync2_7_CDP_Page_CaptureScreenshot()
    Dim ResultCode As Long
    Dim ReceiveCode As Long
    Dim ResponseText As String
    Dim Payload As String
    Dim rid As Long
    Dim ErrorMes As New WinApiError

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync2_1_初期化 を実行してください。"
        Exit Sub
    End If

    rid = CdpNextRequestId()
    Payload = "{""id"":" & CStr(rid) & ",""method"":""Page.captureScreenshot"",""params"":{""format"":""png"",""fromSurface"":true}}"
    ResultCode = g_WebsocketObj.SendMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー(Page.captureScreenshot)。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & rid & ", Page.captureScreenshot) → 受信完了まで待機します"
        ResponseText = WsReceiveUntilNonEmpty(20, ReceiveCode)
        Call CdpDebugPrintReceived(ResponseText, ReceiveCode, ErrorMes)
    End If
End Sub

'* 遷移 → 待機 → スクショまで一気に（各ステップ後に 2_3 で1回ずつ受信）
Sub WebSocketDemoASync2_8_CDP_シナリオ_遷移とスクショ()
    Dim ResultCode As Long
    Dim ResponseText As String
    Dim ErrorMes As New WinApiError
    Dim rid As Long
    Dim Payload As String

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync2_1_初期化 を実行してください。"
        Exit Sub
    End If

    rid = CdpNextRequestId()
    Payload = "{""id"":" & CStr(rid) & ",""method"":""Page.navigate"",""params"":{""url"":""https://example.com/""}}"
    ResultCode = g_WebsocketObj.SendMessage(Payload)
    If ResultCode Then
        Debug.Print "シナリオ: Page.navigate 送信エラー " & ResultCode
        Exit Sub
    End If
    Debug.Print "シナリオ(1/3): Page.navigate 送信 id=" & rid & " → 2_3 相当で1回受信"
    ResponseText = g_WebsocketObj.GetAsyncMessage(, ResultCode)
    Call CdpDebugPrintReceived(ResponseText, ResultCode, ErrorMes)

    On Error Resume Next
    Application.wait (Now + TimeSerial(0, 0, 2))
    On Error GoTo 0

    rid = CdpNextRequestId()
    Payload = "{""id"":" & CStr(rid) & ",""method"":""Page.captureScreenshot"",""params"":{""format"":""png""}}"
    ResultCode = g_WebsocketObj.SendMessage(Payload)
    If ResultCode Then
        Debug.Print "シナリオ: captureScreenshot 送信エラー " & ResultCode
        Exit Sub
    End If
    Debug.Print "シナリオ(2/3): Page.captureScreenshot 送信 id=" & rid & " → 受信完了まで待機"
    ResponseText = WsReceiveUntilNonEmpty(20, ResultCode)
    Call CdpDebugPrintReceived(ResponseText, ResultCode, ErrorMes)

    rid = CdpNextRequestId()
    Payload = "{""id"":" & CStr(rid) & ",""method"":""Network.getAllCookies"",""params"":{}}"
    ResultCode = g_WebsocketObj.SendMessage(Payload)
    If ResultCode Then
        Debug.Print "シナリオ: getAllCookies 送信エラー " & ResultCode
        Exit Sub
    End If
    Debug.Print "シナリオ(3/3): Network.getAllCookies 送信 id=" & rid & " → 1回受信（長文）"
    ResponseText = g_WebsocketObj.GetAsyncMessage(, ResultCode)
    Call CdpDebugPrintReceived(ResponseText, ResultCode, ErrorMes)
End Sub

Private Function CdpJsonEscape(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    CdpJsonEscape = t
End Function

' CDP専用の判定は行わず、非空メッセージが届くまで待つ汎用受信ヘルパー
Private Function WsReceiveUntilNonEmpty(Optional ByVal TimeoutSec As Double = 10, Optional ByRef ResultCode As Long = 0) As String
    Dim tStart As Double
    Dim onceText As String
    Dim ErrorMes As New WinApiError

    tStart = Timer
    ResultCode = 0

    Do
        onceText = g_WebsocketObj.GetAsyncMessage(, ResultCode)
        If ResultCode <> 0 Then Exit Do
        If Len(onceText) > 0 Then
            WsReceiveUntilNonEmpty = onceText
            Exit Function
        End If

        DoEvents
        If Timer - tStart >= TimeoutSec Then Exit Do
    Loop

    If ResultCode <> 0 Then
        Debug.Print "受信エラー。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信タイムアウト(" & TimeoutSec & "秒)。"
    End If
End Function
