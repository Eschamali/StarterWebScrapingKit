Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                   WebSocket のデモンストレーションです（非同期モードのみ）
'
'   SafeTimer 方式のコールバック統一により、AddressOf / isDataReady は不要になりました。
'   コールバックは WebSocketCommunicator.Http_OnCallback で自動的に受け取られます。
'
'   基本的には以下の流れでやり取りが可能です
'   1. ws(s)へ接続
'   2. 受信予約(.RequestWebSocketReceive)を行う
'   3. 何か送ってみる(.SendAsyncMessage)        ※必要に応じて、`.LastSendSuccess`で送信完了したかの待機も可能です
'   4. プロパティメソッド`.LastReceiveExisting`で、データが届いたか確認
'   5. データ取得処理(.GetAsyncMessage)         ※この関数から直接、取得したデータは取れません
'   6. 「.GetAsyncMessage」による実行結果が、0であれば、プロパティメソッド`.LastReceiveContentString`で、4.で得たデータを取得
'   7. 2.からループ
'***************************************************************************************************
Option Explicit



'ポーリング負荷軽減用
Private Declare PtrSafe Sub sleep3 Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)

'変数,オブジェクトの使い回し/保持用に、public化
Private g_WebsocketObj  As WebSocketCommunicator
Private ErrorMes        As New WinApiError
Private SendCount       As Long



'***************************************************************************************************
'                               ■■■ 新規接続用 ■■■
'***************************************************************************************************
'* 機能　　：wss://echo.websocket.org に新規非同期接続し、送受信テストをします
'---------------------------------------------------------------------------------------------------
'* 注意事項：接続後、受信予約を行うとコールバック経由でデータが返ります
'***************************************************************************************************
Sub WebSocketDemoASync_初期化_wss()
    'オブジェクトを作成（SafeTimer 方式のコールバックは Init 内で自動登録される）
    Set g_WebsocketObj = New WebSocketCommunicator
    SendCount = 0

    '接続先を設定します（AddressOf 不要）
    g_WebsocketObj.connectionWebSocket "echo.websocket.org", ""
    Debug.Print "Websocket connect is success. AsyncMode."

    ' 接続直後に受信予約だけ張っておく
    Dim ResultCode As Long
    ResultCode = g_WebsocketObj.RequestWebSocketReceive
    If ResultCode Then
        Debug.Print "受信予約エラー。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信予約結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If

End Sub

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
Sub WebSocketDemoASync_初期化_ws()
    Const CDP_HOST As String = "127.0.0.1"
    Const CDP_PORT As Long = 9222
    Const CDP_TARGET_PATH As String = "devtools/page/065C0DBCE88F6808FEED106660E8D612"
    Dim ResultCode As Long

    ' Chrome は下記のように起動しておく必要があります:
    ' chrome.exe --remote-debugging-port=9222
    ' その後、http://127.0.0.1:9222/json を開いて webSocketDebuggerUrl の page id を確認し、
    ' CDP_TARGET_PATH の REPLACE_WITH_TARGET_ID を置き換えてください。

    Set g_WebsocketObj = New WebSocketCommunicator
    g_WebsocketObj.connectionWebSocket CDP_HOST, CDP_TARGET_PATH, CDP_PORT, False
    SendCount = 0

    Debug.Print "CDP WebSocket connect is success. AsyncMode."

    ' 接続直後に受信予約だけ張っておく
    ResultCode = g_WebsocketObj.RequestWebSocketReceive
    If ResultCode Then
        Debug.Print "受信予約エラー。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信予約結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If
End Sub



'***************************************************************************************************
'                               ■■■ 接続後に行うメソッドDemo ■■■
'***************************************************************************************************
Sub WebSocketDemoASync_受信予約()
    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    '1. 既に予約中か？
    If g_WebsocketObj.isWaitingReceiveResponse Then Debug.Print "既に、受信予約中です": Exit Sub

    '2. 受信予約結果を受け取る
    Dim ResultCode As Long
    ResultCode = g_WebsocketObj.RequestWebSocketReceive

    If ResultCode Then
        g_WebsocketObj.printMsg info_, "受信予約エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"), "Demo"
    Else
        g_WebsocketObj.printMsg info_, "受信予約結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp"), "Demo"
    End If
End Sub

Sub WebSocketDemoASync_受信データを取得()
    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    '1. 受信データが届いてるか？
    If Not g_WebsocketObj.LastReceiveExisting Then Debug.Print "データがまだ、届いてません。": Exit Sub

    '2. 受信メッセージを受け取るようにリクエスト
    Dim ResultCode As Long
    ResultCode = g_WebsocketObj.GetAsyncMessage

    '3. エラーがなければ、受信内容をプロパティメソッドから内容を、取得します
    If ResultCode Then
        g_WebsocketObj.printMsg info_, "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp"), "Demo"
    Else
        g_WebsocketObj.printMsg info_, "受信結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp"), "Demo"
        g_WebsocketObj.printMsg info_, "受信内容：" & g_WebsocketObj.LastReceiveContentString, "Demo"
    End If
End Sub

Sub WebSocketDemoASync_送信()
    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    '1. 送信カウント(任意)
    SendCount = SendCount + 1

    '2. 1件分の送信をしてみる
    Dim ResultCode As Long: ResultCode = g_WebsocketObj.SendAsyncMessage("うみねこ！みゃ～お！" & SendCount & WorksheetFunction.Unichar(129418))

    '3. 送信実行結果
    If ResultCode Then
        Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "送信結果：" & ErrorMes.GetMessage(ResultCode, "WinHttp")
    End If

    '4. 送信がうまくいったかを確認(任意)
    Do
        DoEvents
        sleep3 100
    Loop Until g_WebsocketObj.LastSendSuccess
    Debug.Print "送信がうまくいきました。"
End Sub
