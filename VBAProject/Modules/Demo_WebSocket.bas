Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                          WebSocket のデモンストレーションです（非同期モードのみ）
'
'   SafeTimer 方式のコールバック統一により、AddressOf / isDataReady は不要になりました。
'   コールバックは WebSocketCommunicator.Http_OnCallback で自動的に受け取られます。
'***************************************************************************************************
Option Explicit



'***************************************************************************************************
'                                   ■■■ 非同期処理 ■■■
'***************************************************************************************************
'* 機能　　：wss://echo.websocket.org に新規非同期接続し、送受信テストをします
'---------------------------------------------------------------------------------------------------
'* 注意事項：接続後、受信予約を行うとコールバック経由でデータが返ります
'***************************************************************************************************
Sub WebSocketDemoASync1_1_初期化()
    'オブジェクトを作成（SafeTimer 方式のコールバックは Init 内で自動登録される）
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator

    '接続先を設定します（AddressOf 不要）
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("echo.websocket.org", "")

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success. AsyncMode."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode

        '1件分の送信をしてみる
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("うみねこ！みゃ～お！" & WorksheetFunction.Unichar(129418))

        '実行結果確認
        Dim ErrorMes As New WinApiError
        If ResultCode Then
            Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If

        '受信予約を行う（コールバック経由で Http_OnCallback が発火し m_isReceiving がセットされる）
        Debug.Print WebsocketObj.GetAsyncMessage(, ResultCode)

        If ResultCode Then
            Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub

Sub WebSocketDemoASync1_2_受信リクエスト()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 2172043420336^

    'オブジェクトを作成して、再接続用の LET メソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '受信メッセージを受け取る
    Dim ResultCode As Long
    Debug.Print WebsocketObj.GetAsyncMessage(, ResultCode)

    Dim ErrorMes As New WinApiError
    If ResultCode Then
        Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If
End Sub

Sub WebSocketDemoASync1_3_ハンドルから送信()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 2172043420336^

    'カウント用
    Static Count As Long
    Count = Count + 1

    'オブジェクトを作成して、再接続用の LET メソッドにセット
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '1件分の送信をしてみる
    Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("うみねこ！みゃ～お！" & Count & WorksheetFunction.Unichar(129418))

    Dim ErrorMes As New WinApiError
    If ResultCode Then
        Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If
End Sub

Sub WebSocketDemoASync1_4_後始末()
    '前項で得たハンドル値
    Const ReConnectionHandle As LongPtr = 2519160849248^

    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator
    WebsocketObj.ReConnect = ReConnectionHandle

    '後始末
    WebsocketObj.CloseWebSocket (True)
End Sub

'***************************************************************************************************
'* 機能　　：CDP 経由の長文レスポンスを非同期受信するデモ
'---------------------------------------------------------------------------------------------------
'* 詳細説明：Chrome DevTools Protocol 操作をデモンストレーションします
'***************************************************************************************************
Sub WebSocketDemoASync2_長文レスポンス()
    'オブジェクトを作成
    Dim WebsocketObj As WebSocketCommunicator: Set WebsocketObj = New WebSocketCommunicator

    '接続先を設定します（AddressOf 不要）
    Dim ResultHandleCode As LongPtr: ResultHandleCode = WebsocketObj.Init("127.0.0.1", "devtools/page/1AAA01F8A73F5568DDF8FF042B62D61C", 9222, False)

    '成功判定
    If ResultHandleCode Then
        Debug.Print "Websocket connect is success. AsyncMode."
        Debug.Print "再接続時のハンドルコード：" & ResultHandleCode

        '送信テスト
        Dim ResultCode As Long: ResultCode = WebsocketObj.SendMessage("{""id"":" & 1 & "," & _
                  """method"":""Network.getAllCookies""," & _
                  """params"":{}}")

        Dim ErrorMes As New WinApiError
        If ResultCode Then
            Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If

        '長文受信メッセージを受け取る
        Debug.Print WebsocketObj.GetAsyncMessage(, ResultCode)

        If ResultCode Then
            Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
        Else
            Debug.Print ErrorMes.GetMessage(ResultCode, "WinHttp")
        End If
    Else
        Debug.Print "Websocket connect is failed."
    End If
End Sub
