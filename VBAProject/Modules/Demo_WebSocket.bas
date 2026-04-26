Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                   WebSocket のデモンストレーションです（非同期モードのみ）
'
'   SafeTimer 方式のコールバック統一により、AddressOf / isDataReady は不要になりました。
'   コールバックは WebSocketCommunicator.Http_OnCallback で自動的に受け取られます。
'***************************************************************************************************
Option Explicit

Private g_WebsocketObj As WebSocketCommunicator



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
    Debug.Print "受信内容：" & g_WebsocketObj.GetAsyncMessage(, ResultCode)

    Dim ErrorMes As New WinApiError
    If ResultCode Then
        Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
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
'* 機能　　：CDP 経由の長文レスポンスを非同期受信するデモ
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Chrome DevTools Protocol 操作をデモンストレーションします
'***************************************************************************************************
Sub WebSocketDemoASync2_1_初期化()
    'オブジェクトを作成（SafeTimer 方式のコールバックは Init 内で自動登録される）
    Set g_WebsocketObj = New WebSocketCommunicator

    '接続先を設定します（AddressOf 不要）
    Dim ResultHandleCode As LongPtr: ResultHandleCode = g_WebsocketObj.Init("127.0.0.1", "devtools/page/6125BBE810D35B2F7372345B68E7E653", 9222, False)

    '成功判定
    If ResultHandleCode Then
        Dim g_ReConnectionHandle As LongPtr
        g_ReConnectionHandle = ResultHandleCode
        Debug.Print "Websocket connect is success. AsyncMode."
        Debug.Print "ハンドルコード：" & ResultHandleCode

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
