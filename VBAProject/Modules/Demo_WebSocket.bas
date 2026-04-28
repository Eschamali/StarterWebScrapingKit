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
    g_WebsocketObj.connectionWebSocket "echo.websocket.org"
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
'                             ■■■ 接続後に行う主要メソッドDemo ■■■
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
    Dim timerStart As Single: timerStart = Timer
    Do
        DoEvents
        If Timer - timerStart > 30 Then Err.Raise vbObjectError + 1, , "Timeout waiting for the browser to load the homepage."
    Loop Until g_WebsocketObj.LastSendSuccess
    Debug.Print "送信がうまくいきました。"
End Sub



'***************************************************************************************************
'                         ■■■ DrainPostedCallbacks 判定用（汎用） ■■■
'***************************************************************************************************
'* 目的　　：コールバック排水（Drain 相当）が死んでいないかを、送受信回収率で簡易判定します。いわゆる、ベンチマークテストです
'* 使い方　：
'*   1) WebSocketDemoASync_初期化_wss を実行
'*   2) WebSocketDemoASync_判定_Drain必要性 を実行
'* 判定目安：
'*   - send と recv が概ね一致し、待機状態で詰まらなければ OK
'*   - recv が極端に少ない / isWaitingReceiveResponse が長時間 True 固定なら NG 疑い
'***************************************************************************************************
Sub WebSocketDemoASync_判定_Drain必要性(Optional ByVal BurstCount As Long = 30, Optional ByVal TimeoutSec As Double = 20)
    Const FromProcedureName As String = "Demo_WebSocket.WebSocketDemoASync_判定_Drain必要性"
    Dim i As Long
    Dim sendOk As Long
    Dim recvOk As Long
    Dim rc As Long
    Dim startTick As Double
    Dim msgText As String

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    If BurstCount <= 0 Then BurstCount = 1
    If TimeoutSec <= 0 Then TimeoutSec = 10

    g_WebsocketObj.printMsg info_, "Drain 判定を開始します。BurstCount=" & BurstCount & ", TimeoutSec=" & TimeoutSec, FromProcedureName

    ' 受信予約が未予約なら先に 1 回だけ予約
    If Not g_WebsocketObj.isWaitingReceiveResponse And Not g_WebsocketObj.LastReceiveExisting Then
        rc = g_WebsocketObj.RequestWebSocketReceive
        If rc <> 0 Then
            g_WebsocketObj.printMsg WARN_, "初回の受信予約に失敗しました。ErrorCode=" & rc, FromProcedureName
        End If
    End If

    ' 1) バースト送信
    For i = 1 To BurstCount
        rc = g_WebsocketObj.SendAsyncMessage("[DrainProbe]#" & CStr(i) & "|" & String$(40, "X"))
        If rc = 0 Then
            sendOk = sendOk + 1
        Else
            g_WebsocketObj.printMsg WARN_, "送信失敗 i=" & i & ", ErrorCode=" & rc, FromProcedureName
        End If
        DoEvents
    Next
    g_WebsocketObj.printMsg info_, "送信完了 sendOk=" & sendOk, FromProcedureName

    ' 2) タイムアウトまで受信回収
    startTick = Timer
    Do
        ' 受信データが来ていれば取り出す
        If g_WebsocketObj.LastReceiveExisting Then
            rc = g_WebsocketObj.GetAsyncMessage
            If rc = 0 Then
                recvOk = recvOk + 1
                msgText = g_WebsocketObj.LastReceiveContentString
                g_WebsocketObj.printMsg Debug_, "受信回収 recvOk=" & recvOk & ", Len=" & Len(msgText), FromProcedureName
            Else
                g_WebsocketObj.printMsg WARN_, "GetAsyncMessage 失敗 ErrorCode=" & rc, FromProcedureName
            End If
        End If

        ' 予約が外れていたら再予約
        If Not g_WebsocketObj.isWaitingReceiveResponse And Not g_WebsocketObj.LastReceiveExisting Then
            rc = g_WebsocketObj.RequestWebSocketReceive
            If rc <> 0 Then
                g_WebsocketObj.printMsg WARN_, "再予約失敗 ErrorCode=" & rc, FromProcedureName
            End If
        End If

        If recvOk >= sendOk And sendOk > 0 Then Exit Do
        DoEvents
        sleep3 10
    Loop While Timer - startTick < TimeoutSec

    ' 3) 判定出力
    g_WebsocketObj.printMsg info_, "Drain 判定結果 sendOk=" & sendOk & ", recvOk=" & recvOk & _
                                   ", waiting=" & g_WebsocketObj.isWaitingReceiveResponse & _
                                   ", hasData=" & g_WebsocketObj.LastReceiveExisting, FromProcedureName, True

    If sendOk = 0 Then
        g_WebsocketObj.printMsg WARN_, "送信成功が 0 件のため判定不能です。接続状態を確認してください。", FromProcedureName
    ElseIf recvOk >= sendOk Then
        g_WebsocketObj.printMsg info_, "OK: 排水（Drain 相当）は機能している可能性が高いです。", FromProcedureName
    ElseIf g_WebsocketObj.isWaitingReceiveResponse And Not g_WebsocketObj.LastReceiveExisting Then
        g_WebsocketObj.printMsg WARN_, "NG疑い: 受信待機が詰まり気味です（Drain が機能していない可能性）。", FromProcedureName
    Else
        g_WebsocketObj.printMsg WARN_, "要観察: 回収率が低いです。BurstCount/TimeoutSec を変えて再試験してください。", FromProcedureName
    End If
End Sub



'***************************************************************************************************
'                   ■■■ chrome devtools Protocol 用の簡易コマンド ※送信のみ ■■■
'***************************************************************************************************
Sub WebSocketDemoASync_CDP送信_RuntimeEvaluate()
    Dim ResultCode As Long
    Dim Payload As String

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    SendCount = SendCount + 1
    Payload = "{""id"":" & CStr(SendCount) & ",""method"":""Runtime.evaluate"",""params"":{""expression"":""document.title"",""returnByValue"":true}}"

    ResultCode = g_WebsocketObj.SendAsyncMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & SendCount & ", Runtime.evaluate)"
    End If
End Sub

'* Network.getAllCookies → クッキー一覧の長い JSON（長文レスポンスの負荷テスト向け）
Sub WebSocketDemoASync2_5_CDP_Network_GetAllCookies()
    Dim ResultCode As Long
    Dim Payload As String

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    SendCount = SendCount + 1
    Payload = "{""id"":" & CStr(SendCount) & ",""method"":""Network.getAllCookies"",""params"":{}}"
    ResultCode = g_WebsocketObj.SendAsyncMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー(Network.getAllCookies)。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & SendCount & ", Network.getAllCookies) → 別途受信Demoプロシージャを実行してください"
    End If
End Sub

'* Page.navigate → 実際にタブの URL が変わる
Sub WebSocketDemoASync2_6_CDP_Page_Navigate(Optional ByVal TargetUrl As String)
    Dim ResultCode As Long
    Dim Payload As String
    Dim esc As String
    Dim ErrorMes As New WinApiError

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    If Len(TargetUrl) = 0 Then TargetUrl = "https://www.wikipedia.org/"
    esc = CdpJsonEscape(TargetUrl)

    SendCount = SendCount + 1
    Payload = "{""id"":" & CStr(SendCount) & ",""method"":""Page.navigate"",""params"":{""url"":""" & esc & """}}"
    ResultCode = g_WebsocketObj.SendAsyncMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー(Page.navigate)。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & SendCount & ", Page.navigate url=" & TargetUrl & ") → ブラウザの表示を確認し、必要なら 別途受信Demoプロシージャを実行してください"
    End If
End Sub

'* Page.captureScreenshot → result.data に base64 PNG（長文になりがち）
Sub WebSocketDemoASync2_7_CDP_Page_CaptureScreenshot()
    Dim ResultCode As Long
    Dim ReceiveCode As Long
    Dim ResponseText As String
    Dim Payload As String
    Dim ErrorMes As New WinApiError

    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    SendCount = SendCount + 1
    Payload = "{""id"":" & CStr(SendCount) & ",""method"":""Page.captureScreenshot"",""params"":{""format"":""png"",""fromSurface"":true}}"
    ResultCode = g_WebsocketObj.SendAsyncMessage(Payload)
    If ResultCode Then
        Debug.Print "CDP送信エラー(Page.captureScreenshot)。ErrorCode：" & ResultCode & ",Description：" & ErrorMes.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "CDP送信OK(id=" & SendCount & ", Page.captureScreenshot) → 別途受信Demoプロシージャを実行してください"
    End If
End Sub


Private Function CdpJsonEscape(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    CdpJsonEscape = t
End Function
