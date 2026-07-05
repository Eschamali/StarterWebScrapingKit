Attribute VB_Name = "Demo_WebSocket"
'***************************************************************************************************
'                   WebSocket のデモンストレーションです（非同期モードのみ）
'
'   基本的には以下の流れでやり取りが可能です
'   1. ws(s)へ接続
'   2. 受信予約(.RequestWebSocketReceive)を行う
'   3. 何か送ってみる(.SendAsyncMessageAsUTF8/BINARY)   ※必要に応じて、`.LastSendSuccess`で送信完了したかの待機も可能です
'   4. プロパティメソッド`.LastReceiveExisting`で、データが届いたか確認
'   5. データ取得処理(.GetAsyncMessage)                 ※この関数から直接、取得したデータは取れません
'   6. 「.GetAsyncMessage」による実行結果が、0であれば、プロパティメソッド`.LastReceiveContentUTF8/BINARY`で、4.で得たデータを取得
'   7. 2.からループ
'***************************************************************************************************
Option Explicit
Option Private Module



'ポーリング負荷軽減用
Private Declare PtrSafe Sub sleep3 Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)

'変数,オブジェクトの使い回し/保持用に、public化
Private g_WebsocketObj  As WebSocketCommunicator
Private SendCount       As Long
Private wsForChromiumobj As WebSocketCommunicator



'***************************************************************************************************
'* ★ ワンアクションの「安全リセット」マクロ ★
'*   下記の3ステップを1アクションでまとめ、VBE のリセットボタンの代替として使うためのものです。
'*     ① WinHttp の status callback を解除して、worker thread を静止
'*     ② SetWindowLongPtr で差し替えたウィンドウフックを解除し、元の WndProc を全件復元
'*     ③ VBA の `End` ステートメントで状態をリセット (= リセットボタン押下相当)
'*
'*   ①②を踏まえてから③を実行することで、`0xc0000027 (STATUS_BAD_FUNCTION_TABLE)` 系の
'*   Excel クラッシュを大幅に回避できます。VBE のリセットボタンの代わりにこちらをご利用ください。
'*
'* ◆ 使い方
'*     A. イミディエイトに `Demo_SafeReset` と入力して Enter
'*     B. Excel のクイックアクセスツールバーに登録 → ボタン1クリックで完了
'*        (ファイル → オプション → クイックアクセスツールバー → "マクロ" → Demo_SafeReset を追加)
'*     C. ショートカット割当：別の Workbook_Open 等で
'*           Application.OnKey "^+r", "Demo_SafeReset"   ' Ctrl+Shift+R で発火
'*        を実行しておけば、ホットキー1発で発火させることも可能です。
'*
'* ◆ 注意事項
'*     - 実行後は VBA の状態が完全に初期化されるため、続けて使う場合は初期化マクロから再開してください。
'*     - VBE のブレーク中 (黄色い行で停止) でもイミディエイトから実行可能です。
'*       ただし `WebSocket_OnCallback` 内で停止中に限り、COM event 経由のスタックを `End` でも
'*       巻き戻しきれずクラッシュする場合があります。その場合は F8 等で当該 Sub を抜けてから実行してください。
'***************************************************************************************************
Public Sub Demo_SafeReset()
    '対象オブジェクトを指定
    Dim TargetClean As WebSocketCommunicator: Set TargetClean = g_WebsocketObj


    '①② callback 解除 + WndProc 復元 (= 既存の手動掃除と同じ)
    On Error Resume Next
    If Not g_WebsocketObj Is Nothing Then TargetClean.EmergencyUnregisterWinHttpCallbacks    '※ここは、使用している`WebSocketCommunicator`Classオブジェクト名に合わせること
    RemoveWinHttpMessageHook

    '念のため Class_Terminate を走らせて、内部 handle (HINTERNET / hWebsocket) を確実に close
    Set TargetClean = Nothing
    On Error GoTo 0

    '③ VBA の状態を全てクリア (= リセット相当の制御された unwind)
    End
End Sub



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
        Debug.Print "受信予約エラー。ErrorCode：" & ResultCode & ",Description：" & WinApiError.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信予約結果：" & WinApiError.GetMessage(ResultCode, "WinHttp")
    End If

End Sub

'***************************************************************************************************
'* 機能　　：CDP（Chrome DevTools Protocol）を WebSocket 経由で叩くデモ
'---------------------------------------------------------------------------------------------------
'* 詳細説明：1. `--remote-debugging-Port=9222`付きでChromium起動か、「edge://inspect/#remote-debugging」で「Allow remote debugging for this browser instance」を有効化
'            2. 「http://127.0.0.1:9222/json/version」にアクセスか、Environ("UserProfile") & "\AppData\Local\Microsoft\Edge\User Data\DevToolsActivePort"で、接続先WebSocketURLを特定
'            3. このプロシージャを実行。ブラウザから案内が出たら「許可」を選択
'***************************************************************************************************
Sub WebSocketModeForCDP()
    Const CDP_HOST As String = "127.0.0.1"
    Const CDP_PORT As Long = 9222
    Const CDP_TARGET_PATH As String = "/devtools/browser/f7a90a36-75b5-4fb7-90dc-c8b871f6cbe2"
    Dim ResultCode As Long

    ' Chrome は下記のように起動しておく必要があります:
    ' chrome.exe --remote-debugging-port=9222
    ' その後、http://127.0.0.1:9222/json を開いて webSocketDebuggerUrl の page id を確認し、
    ' CDP_TARGET_PATH の REPLACE_WITH_TARGET_ID を置き換えてください。

    Set wsForChromiumobj = New WebSocketCommunicator
    wsForChromiumobj.connectionWebSocket CDP_HOST, CDP_TARGET_PATH, CDP_PORT, False
    SendCount = 0

    Debug.Print "CDP WebSocket connect is success. AsyncMode."

    ' 接続直後に受信予約だけ張っておく
    ResultCode = wsForChromiumobj.RequestWebSocketReceive
    If ResultCode Then
        Debug.Print "受信予約エラー。ErrorCode：" & ResultCode & ",Description：" & WinApiError.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信予約結果：" & WinApiError.GetMessage(ResultCode, "WinHttp")
    End If

    '設定セルから、ユーザ名を取得
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_WebSocket.WebSocketModeForCDP")

    'データ枠のみ確保 ※Pipe版ロジックとの互換性を保つため
    Dim hoge1 As New CDPCore: hoge1.serialize UserName
    '1. 必要なデータを`Dictionary`に詰める
    Dim BrowserInfo As New Dictionary
    BrowserInfo.Add "BiDi-context", vbNullString
    BrowserInfo.Add "sessionID", vbNullString
    BrowserInfo.Add "targetID", vbNullString

    '2. Excelのテーブルへ記録する
    Set ShSetting01_StartBrowser.TableBrowserContext(UserName, "Demo_WebSocket.WebSocketModeForCDP") = BrowserInfo

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
        Debug.Print "受信予約エラー発生。ErrorCode：" & ResultCode & ",Description：" & WinApiError.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信予約結果：" & WinApiError.GetMessage(ResultCode, "WinHttp")
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
        Debug.Print "受信エラー発生。ErrorCode：" & ResultCode & ",Description：" & WinApiError.GetMessage(ResultCode, "winhttp")
    Else
        Debug.Print "受信結果：" & WinApiError.GetMessage(ResultCode, "WinHttp")
        Debug.Print "受信内容：" & g_WebsocketObj.LastReceiveContentUTF8
    End If
End Sub

Sub WebSocketDemoASync_送信()
    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    '1. 送信カウント(任意)
    SendCount = SendCount + 1

    '2. 1件分の送信をしてみる(`WorksheetFunction.Unichar`で、絵文字送信も可能)
    Dim ResultCode As Long: ResultCode = g_WebsocketObj.SendAsyncMessageAsUTF8("うみねこ！みゃ～お！" & SendCount & WorksheetFunction.Unichar(129418))

    '3. 送信実行結果
    If ResultCode Then
        Debug.Print "送信エラー発生。ErrorCode：" & ResultCode & ",Description：" & WinApiError.GetMessage(ResultCode, "winhttp")
        Exit Sub
    Else
        Debug.Print "送信結果：" & WinApiError.GetMessage(ResultCode, "WinHttp")
    End If

    '4. 送信がうまくいったかを確認(任意)
    Dim timerStart As Double: timerStart = g_WebsocketObj.TimerCounter
    Do Until g_WebsocketObj.LastSendSuccess
        DoEvents
        If g_WebsocketObj.TimerCounter - timerStart > 30000 Then Err.Raise vbObjectError + 1, , "Timeout waiting for the WebSocket to send result."
    Loop
    Debug.Print "送信がうまくいきました。"
End Sub

Sub WebSocketDemo_Close()
    If g_WebsocketObj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws/wss を実行してください。"
        Exit Sub
    End If

    g_WebsocketObj.CloseWebSocket
    Set g_WebsocketObj = Nothing
    Debug.Print "WebSocketを閉じました"
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
Sub WebSocketDemoASync_判定_Drain必要性(Optional ByVal BurstCount As Long = 30, Optional ByVal TimeoutMSec As Double = 20000)
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
    If TimeoutMSec <= 0 Then TimeoutMSec = 10000

    g_WebsocketObj.printMsg info_, "Drain 判定を開始します。BurstCount=" & BurstCount & ", TimeoutSec=" & TimeoutMSec, FromProcedureName

    ' 受信予約が未予約なら先に 1 回だけ予約
    If Not g_WebsocketObj.isWaitingReceiveResponse And Not g_WebsocketObj.LastReceiveExisting Then
        rc = g_WebsocketObj.RequestWebSocketReceive
        If rc <> 0 Then
            g_WebsocketObj.printMsg WARN_, "初回の受信予約に失敗しました。ErrorCode=" & rc, FromProcedureName
        End If
    End If

    ' 1) バースト送信
    For i = 1 To BurstCount
        rc = g_WebsocketObj.SendAsyncMessageAsUTF8("[DrainProbe]#" & CStr(i) & "|" & String$(40, "X"))
        If rc = 0 Then
            sendOk = sendOk + 1
        Else
            g_WebsocketObj.printMsg WARN_, "送信失敗 i=" & i & ", ErrorCode=" & rc, FromProcedureName
        End If
        DoEvents
    Next
    g_WebsocketObj.printMsg info_, "送信完了 sendOk=" & sendOk, FromProcedureName

    ' 2) タイムアウトまで受信回収
    startTick = g_WebsocketObj.TimerCounter
    Do
        ' 受信データが来ていれば取り出す
        If g_WebsocketObj.LastReceiveExisting Then
            rc = g_WebsocketObj.GetAsyncMessage
            If rc = 0 Then
                recvOk = recvOk + 1
                msgText = g_WebsocketObj.LastReceiveContentUTF8
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
    Loop While g_WebsocketObj.TimerCounter - startTick < TimeoutMSec

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
'                           ■■■ WebSocket経由でのCDP制御Demo ■■■
'***************************************************************************************************
Sub WebSocketDemoASync_CDP送信_PageNavigate()
    If wsForChromiumobj Is Nothing Then
        Debug.Print "先に WebSocketModeForCDP を実行してください。"
        Exit Sub
    End If


    '設定セルから、ユーザ名を取得
    Dim c As New CDPBrowser
    Dim r As CDPContext
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_WebSocket.WebSocketDemoASync_CDP送信_PageNavigate")

    '1. Excelに記録されてるパイプハンドル情報の生存確認
    If Not c.reattach(UserName, wsForChromiumobj) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub

    '2. 未接続のタブに接続
    '※この時、必ず`setMain:=True`とすること。必要に応じて検索条件(URLマッチ等)も設定して下さい
    Set r = c.getTab(setMain:=True)
'    Set r = c.newTab(setMain:=True) '新しいタブ生成からでもOK


    '3．別ページに遷移して終了
    r.navigate "https://kemono-friends.jp/"
End Sub

'* Network.getAllCookies → クッキー一覧の長い JSON（長文レスポンスの負荷テスト向け）
Sub WebSocketDemoASync2_5_CDP_Network_GetAllCookies()
    If wsForChromiumobj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws を実行してください。"
        Exit Sub
    End If


    '設定セルから、ユーザ名を取得
    Dim c As New CDPBrowser
    Dim r As CDPContext
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_WebSocket.WebSocketDemoASync2_5_CDP_Network_GetAllCookies")

    '1. Excelに記録されてるパイプハンドル情報の生存確認
    If Not c.reattach(UserName, wsForChromiumobj) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub

    '2. 未接続のタブに接続
    '※この時、必ず`setMain:=True`とすること。必要に応じて検索条件(URLマッチ等)も設定して下さい
    Set r = c.getTab(setMain:=True)
'    Set r = c.newTab(setMain:=True) '新しいタブ生成からでもOK


    Debug.Print "接続先のブラウザから、cookieを出力します..." & vbCrLf & r.ExecuteCDP("Network.getAllCookies").Stringify
End Sub

'* Page.captureScreenshot 保存
Sub WebSocketDemoASync2_7_CDP_Page_CaptureScreenshot()
    If wsForChromiumobj Is Nothing Then
        Debug.Print "先に WebSocketDemoASync_初期化_ws を実行してください。"
        Exit Sub
    End If


    '設定セルから、ユーザ名を取得
    Dim c As New CDPBrowser
    Dim r As CDPContext
    Dim UserName As String
    UserName = ShSetting01_StartBrowser.UseRangeID(2, "Demo_WebSocket.WebSocketDemoASync2_7_CDP_Page_CaptureScreenshot")

    '1. Excelに記録されてるパイプハンドル情報の生存確認
    If Not c.reattach(UserName, wsForChromiumobj) Then MsgBox "「" & UserName & "」に接続できませんでした。WebSocket情報がお亡くなりです。", vbCritical, "Chrome DevTools Protocol": Exit Sub

    '2. 未接続のタブに接続
    '※この時、必ず`setMain:=True`とすること。必要に応じて検索条件(URLマッチ等)も設定して下さい
    Set r = c.getTab(setMain:=True)
'    Set r = c.newTab(setMain:=True) '新しいタブ生成からでもOK


    r.snapPage Environ("UserProfile") & "\Downloads", "test.png"
    r.notify "WebSocket経由で、スクショ保存しました" & WorksheetFunction.Unichar(129418)
End Sub
